from collections.abc import Generator
from typing import Any
import tempfile
import os
import zipfile
import xml.etree.ElementTree as ET
import io

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from markitdown import MarkItDown

# ---------------------------------------------------------------------------
# Namespaces DrawingML / WordprocessingML
# ---------------------------------------------------------------------------
_NS = {
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _parse_pct(val: str) -> float:
    """Convertit une valeur srcRect Word (perMille×1000) en float 0..1.
    Les valeurs négatives (zoom-out Word) sont clampées à 0."""
    return max(0.0, int(val) / 100_000.0)


def _has_real_crop(attrs: dict) -> bool:
    """Retourne True si au moins une valeur est positive (crop réel)."""
    return any(int(v) > 0 for v in attrs.values())


def _crop_image_bytes(img_bytes: bytes, attrs: dict) -> bytes | None:
    """Applique le crop et retourne les bytes, ou None si crop nul/invalide."""
    from PIL import Image
    img = Image.open(io.BytesIO(img_bytes))
    fmt = img.format or "PNG"
    w, h = img.size

    l = _parse_pct(attrs.get("l", "0"))
    t = _parse_pct(attrs.get("t", "0"))
    r = _parse_pct(attrs.get("r", "0"))
    b = _parse_pct(attrs.get("b", "0"))

    left, top     = max(0, int(w * l)), max(0, int(h * t))
    right, bottom = min(w, int(w * (1.0 - r))), min(h, int(h * (1.0 - b)))

    if left >= right or top >= bottom or (left, top, right, bottom) == (0, 0, w, h):
        return None

    out = io.BytesIO()
    save_fmt = fmt if fmt in ("PNG", "JPEG", "GIF", "BMP", "TIFF") else "PNG"
    img.crop((left, top, right, bottom)).save(out, format=save_fmt)
    return out.getvalue()


def _apply_docx_crops(src_path: str, dst_path: str) -> bool:
    """
    Lit src_path (.docx), applique physiquement les crops Word (a:srcRect).

    Corrections vs version naïve :
    - Images partagées avec crops différents (sprite sheets) : crée une nouvelle
      image par usage et met à jour les relations, au lieu d'écraser le fichier.
    - Valeurs srcRect négatives (zoom-out Word) : ignorées (clampées à 0).
    - srcRect vide : ignoré.

    Retourne True si au moins un crop a été appliqué.
    """
    try:
        from PIL import Image  # noqa: F401 — vérification import
    except ImportError:
        return False

    import tempfile as _tempfile
    from pathlib import Path

    with _tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)

        with zipfile.ZipFile(src_path, "r") as z:
            z.extractall(tmp)

        rels_path = tmp / "word" / "_rels" / "document.xml.rels"
        if not rels_path.exists():
            return False

        rels_tree = ET.parse(rels_path)
        rels_root = rels_tree.getroot()
        rel_map = {r.get("Id"): r.get("Target") for r in rels_root}

        # Prochain rId disponible
        existing_ids = [int(r.get("Id", "rId0").replace("rId", ""))
                        for r in rels_root if (r.get("Id") or "").startswith("rId")]
        _next_rid = [max(existing_ids, default=0) + 1]

        def _new_rid():
            rid = f"rId{_next_rid[0]}"
            _next_rid[0] += 1
            return rid

        doc_path = tmp / "word" / "document.xml"
        doc_tree = ET.parse(doc_path)
        doc_root = doc_tree.getroot()

        crops_applied = 0

        for blipFill in doc_root.iter(f"{{{_NS['pic']}}}blipFill"):
            src_rect = blipFill.find(f"{{{_NS['a']}}}srcRect")
            if src_rect is None:
                continue

            attrs = {k: src_rect.get(k, "0")
                     for k in ("l", "t", "r", "b") if src_rect.get(k) is not None}

            if not attrs or not _has_real_crop(attrs):
                continue  # pas de crop réel

            blip = blipFill.find(f"{{{_NS['a']}}}blip")
            if blip is None:
                continue

            embed = blip.get(f"{{{_NS['r']}}}embed")
            if embed not in rel_map:
                continue

            img_path = tmp / "word" / rel_map[embed]
            if not img_path.exists():
                continue

            try:
                cropped_bytes = _crop_image_bytes(img_path.read_bytes(), attrs)
            except Exception:
                continue

            if cropped_bytes is None:
                continue

            # Créer une nouvelle image pour ce crop (safe même si image non partagée)
            new_name = f"image_crop_{crops_applied + 1}{img_path.suffix}"
            new_path = tmp / "word" / "media" / new_name
            new_path.write_bytes(cropped_bytes)

            # Nouvelle relation
            rid = _new_rid()
            rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            new_rel = ET.SubElement(rels_root, "Relationship")
            new_rel.set("Id", rid)
            new_rel.set("Type", rel_ns)
            new_rel.set("Target", f"media/{new_name}")

            # Mettre à jour le blip et supprimer srcRect
            blip.set(f"{{{_NS['r']}}}embed", rid)
            blipFill.remove(src_rect)
            crops_applied += 1

        if not crops_applied:
            return False

        doc_tree.write(doc_path, xml_declaration=True, encoding="UTF-8")
        rels_tree.write(rels_path, xml_declaration=True, encoding="UTF-8")

        with zipfile.ZipFile(dst_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for f in tmp.rglob("*"):
                if f.is_file():
                    zout.write(f, f.relative_to(tmp))

    return True


# ---------------------------------------------------------------------------
# Dify Tool
# ---------------------------------------------------------------------------

class MarkitdownTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage, None, None]:
        files = tool_parameters.get('files', [])

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

        for file in files:
            try:
                file_extension = file.extension if file.extension else '.tmp'

                with tempfile.NamedTemporaryFile(delete=False, suffix=file_extension) as temp_file:
                    temp_file.write(file.blob)
                    temp_file_path = temp_file.name

                cropped_path = None
                try:
                    convert_path = temp_file_path

                    # Appliquer les crops avant conversion pour les .docx
                    if file_extension.lower() in ('.docx', '.docm'):
                        cropped_path = temp_file_path + "_cropped.docx"
                        if _apply_docx_crops(temp_file_path, cropped_path):
                            convert_path = cropped_path

                    md = MarkItDown()
                    result = md.convert(convert_path)

                    yield self.create_blob_message(
                        result.text_content.encode(),
                        meta={"mime_type": "text/markdown"},
                    )

                    if result and hasattr(result, 'text_content'):
                        results.append({"filename": file.filename, "content": result.text_content})
                        json_results.append({
                            "filename": file.filename,
                            "original_format": file_extension.lstrip('.'),
                            "markdown_content": result.text_content,
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

        yield self.create_json_message({
            "status": "success" if results else "error",
            "total_files": len(files),
            "successful_conversions": len(results),
            "results": json_results
        })

        if not results:
            yield self.create_text_message("No files were successfully processed")
        elif len(results) == 1:
            yield self.create_text_message(results[0]["content"])
        else:
            combined = ""
            for idx, r in enumerate(results, 1):
                combined += f"\n{'='*50}\nFile {idx}: {r['filename']}\n{'='*50}\n\n"
                combined += r["content"] + "\n\n"
            yield self.create_text_message(combined.strip())
