from collections.abc import Generator
from typing import Any
import tempfile
import os
import re
import base64
import hashlib
import requests
from urllib.parse import urlparse, urlunparse

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from markitdown import MarkItDown
from tools.apply_docx_crops import apply_docx_crops

class MarkitdownTool(Tool):

    def _upload_image(self, filename: str, content: bytes, mimetype: str) -> dict:
        """
        Upload an image via the Dify backwards invocation, rewriting the
        signed URL host if FILES_URL is set (needed for local dev where
        'ai-dify-api' is not resolvable from the host machine).
        """
        from dify_plugin.core.entities.invocation import InvokeType

        for upload_data in self.session.file._backwards_invoke(
            InvokeType.UploadFile,
            dict,
            {
                "filename": filename,
                "mimetype": mimetype,
            },
        ):
            signed_url = upload_data.get("url")
            if not signed_url:
                raise Exception("upload file failed, could not get signed url")

            # Rewrite host if FILES_URL is set
            files_url = os.environ.get("FILES_URL")
            if files_url:
                parsed = urlparse(signed_url)
                target = urlparse(files_url)
                signed_url = urlunparse(
                    parsed._replace(scheme=target.scheme, netloc=target.netloc)
                )

            response = requests.post(
                signed_url,
                files={"file": (filename, content, mimetype)},
                timeout=30,
            )
            if response.status_code != 201:
                raise Exception(
                    f"upload file failed, status: {response.status_code}, "
                    f"response: {response.text}"
                )
            result = response.json()
            # Rewrite preview_url host if FILES_URL is set
            preview_url = result.get("preview_url")
            if preview_url and files_url:
                parsed_preview = urlparse(preview_url)
                target = urlparse(files_url)
                result["preview_url"] = urlunparse(
                    parsed_preview._replace(scheme=target.scheme, netloc=target.netloc)
                )
            return result

        raise Exception("upload file failed, empty response from server")

    def _replace_base64_images_with_urls(self, markdown: str) -> str:
        """
        Find all base64 data URIs in markdown image tags, upload each image
        via the Dify file API, and replace the data URI with the signed
        preview_url returned by Dify.

        :param markdown: markdown text potentially containing base64 images
        :return: markdown with data URIs replaced by permanent signed URLs
        """
        # Match ![alt](data:image/TYPE;base64,DATA)
        pattern = re.compile(
            r'(!\[[^\]]*\])\((data:([^;]+);base64,([^)]+))\)'
        )

        # Deduplicate: cache hash -> URL to avoid re-uploading identical images
        hash_to_url: dict[str, str] = {}

        def replace_match(m: re.Match) -> str:
            alt_part = m.group(1)        # e.g. ![alt text]
            mime_type = m.group(3)       # e.g. image/png
            b64_data  = m.group(4)       # raw base64 string

            try:
                image_bytes = base64.b64decode(b64_data)
            except Exception:
                # Leave malformed data URIs untouched
                return m.group(0)

            # Stable filename based on content hash
            content_hash = hashlib.sha256(image_bytes).hexdigest()[:16]
            ext = mime_type.split('/')[-1].split('+')[0]  # png, jpeg, webp …
            filename = f"{content_hash}.{ext}"

            if content_hash in hash_to_url:
                url = hash_to_url[content_hash]
            else:
                try:
                    uploaded = self._upload_image(filename, image_bytes, mime_type)
                    # preview_url is a signed /files/tools/{uuid}.ext URL
                    url = uploaded.get("preview_url")
                    if not url:
                        raise Exception("no preview_url in upload response")
                    hash_to_url[content_hash] = url
                except Exception:
                    # If upload fails, keep the original data URI
                    return m.group(0)

            return f"{alt_part}({url})"

        return pattern.sub(replace_match, markdown)

    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage, None, None]:
        files = tool_parameters.get('files', [])
        enable_base64_images = tool_parameters.get('enable_base64_images', False)
        upload_images = tool_parameters.get('upload_images', False)
        
        # Handle empty files array
        if not files:
            yield self.create_text_message("No files provided")
            yield self.create_json_message({
                "status": "error",
                "message": "No files provided",
                "results": []
            })
            return

        results = []
        json_results = []
        
        # Process each file
        for file in files:
            try:
                file_extension = file.extension if file.extension else '.tmp'
                
                # Rewrite file URL if FILES_URL is set (useful for local debug)
                files_url = os.environ.get("FILES_URL")
                if files_url:
                    parsed = urlparse(file.url)
                    target = urlparse(files_url)
                    rewritten = parsed._replace(scheme=target.scheme, netloc=target.netloc)
                    file.url = urlunparse(rewritten)

                with tempfile.NamedTemporaryFile(delete=False, suffix=file_extension) as temp_file:
                    temp_file.write(file.blob)
                    temp_file_path = temp_file.name
                
                cropped_path = None
                try:
                    # Appliquer les crops DOCX avant conversion
                    convert_path = temp_file_path
                    if file_extension.lower() in (".docx", ".docm"):
                        cropped_path = temp_file_path + "_cropped" + file_extension
                        if apply_docx_crops(temp_file_path, cropped_path):
                            convert_path = cropped_path

                    md = MarkItDown()
                    # Keep data URIs if base64 output is requested OR if we
                    # plan to upload them (we need the raw bytes to upload)
                    keep_uris = enable_base64_images or upload_images
                    result = md.convert(convert_path, keep_data_uris=keep_uris)

                    markdown_content = result.text_content

                    # Replace base64 images with uploaded Dify file URLs
                    if upload_images and markdown_content:
                        markdown_content = self._replace_base64_images_with_urls(
                            markdown_content
                        )

                    # Create blob message for backward compatibility
                    yield self.create_blob_message(
                        markdown_content.encode(),
                        meta={
                            "mime_type": "text/markdown",
                        },
                    )
                    
                    if result and hasattr(result, 'text_content'):
                        results.append({
                            "filename": file.filename,
                            "content": markdown_content
                        })
                        
                        # Add to JSON results
                        json_results.append({
                            "filename": file.filename,
                            "original_format": file_extension.lstrip('.'),
                            "markdown_content": markdown_content,
                            "status": "success"
                        })
                    else:
                        error_msg = f"Conversion failed for file {file.filename}. Result: {result}"
                        yield self.create_text_message(text=error_msg)
                        json_results.append({
                            "filename": file.filename,
                            "original_format": file_extension.lstrip('.'),
                            "error": error_msg,
                            "status": "error"
                        })
                        
                finally:
                    for path in [temp_file_path, cropped_path]:
                        if path and os.path.exists(path):
                            os.unlink(path)
                        
            except Exception as e:
                error_msg = f"Error processing file {file.filename}: {str(e)}"
                yield self.create_text_message(text=error_msg)
                json_results.append({
                    "filename": file.filename,
                    "original_format": file_extension.lstrip('.'),
                    "error": error_msg,
                    "status": "error"
                })
        
        # Create JSON response
        json_response = {
            "status": "success" if len(results) > 0 else "error",
            "total_files": len(files),
            "successful_conversions": len(results),
            "results": json_results
        }
        yield self.create_json_message(json_response)
        
        # Return text results based on number of files processed (for backward compatibility)
        if len(results) == 0:
            yield self.create_text_message("No files were successfully processed")
        elif len(results) == 1:
            yield self.create_text_message(results[0]["content"])
        else:
            combined_content = ""
            for idx, result in enumerate(results, 1):
                combined_content += f"\n{'='*50}\n"
                combined_content += f"File {idx}: {result['filename']}\n"
                combined_content += f"{'='*50}\n\n"
                combined_content += result['content']
                combined_content += "\n\n"
            
            yield self.create_text_message(combined_content.strip())
