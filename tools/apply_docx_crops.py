from collections.abc import Generator
from typing import Any
import tempfile
import os
import zipfile
import xml.etree.ElementTree as ET

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

# ---------------------------------------------------------------------------
# Namespaces DrawingML / WordprocessingML
# ---------------------------------------------------------------------------
_NS = {
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _parse_crop_pct(val: str) -> float:
    """Word stocke les valeurs srcRect en 1/100 000 (ex: 10000 = 10%)."""
    return int(val) / 100_000.0


def apply_docx_crops(src_path: str, dst_path: str) -> bool:
    """
    Lit src_path (.docx), applique physiquement les crops Word (a:srcRect)
    sur chaque image avec Pillow, supprime le srcRect du XML, et écrit dst_path.
    Retourne True si au moins une image a été croppée.
    """
    try:
        from PIL import Image
    except ImportError:
        return False

    modified = False

    with tempfile.TemporaryDirectory() as tmp:
        # 1. Extraire le docx
        with zipfile.ZipFile(src_path, "r") as z:
            z.extractall(tmp)

        # 2. Lire les relations
        rels_path = os.path.join(tmp, "word", "_rels", "document.xml.rels")
        if not os.path.exists(rels_path):
            return False

        rels_root = ET.parse(rels_path).getroot()
        rel_map = {r.get("Id"): r.get("Target") for r in rels_root}

        # 3. Parser document.xml
        doc_path = os.path.join(tmp, "word", "document.xml")
        doc_tree = ET.parse(doc_path)
        doc_root = doc_tree.getroot()

        # 4. Parcourir les blipFill avec srcRect
        for blipFill in doc_root.iter(f"{{{_NS['pic']}}}blipFill"):
            src_rect = blipFill.find(f"{{{_NS['a']}}}srcRect")
            if src_rect is None:
                continue

            attrs = {k: src_rect.get(k, "0") for k in ("l", "t", "r", "b")}
            if all(v == "0" for v in attrs.values()):
                continue

            blip = blipFill.find(f"{{{_NS['a']}}}blip")
            if blip is None:
                continue

            embed = blip.get(f"{{{_NS['r']}}}embed")
            if embed not in rel_map:
                continue

            img_rel = rel_map[embed]
            # Les chemins de relations sont relatifs à word/
            img_abs = os.path.normpath(os.path.join(tmp, "word", img_rel))
            if not os.path.exists(img_abs):
                continue

            try:
                img = Image.open(img_abs)
                fmt = img.format or "PNG"
                w, h = img.size

                l = _parse_crop_pct(attrs["l"])
                t = _parse_crop_pct(attrs["t"])
                r = _parse_crop_pct(attrs["r"])
                b = _parse_crop_pct(attrs["b"])

                left   = int(w * l)
                top    = int(h * t)
                right  = int(w * (1.0 - r))
                bottom = int(h * (1.0 - b))

                if left >= right or top >= bottom:
                    continue

                cropped = img.crop((left, top, right, bottom))
                save_fmt = fmt if fmt in ("PNG", "JPEG", "GIF", "BMP", "TIFF") else "PNG"
                cropped.save(img_abs, format=save_fmt)

                # Supprimer srcRect pour éviter un double-crop
                blipFill.remove(src_rect)
                modified = True

            except Exception:
                continue  # on garde l'image originale si erreur

        if not modified:
            return False

        # 5. Réécrire document.xml
        doc_tree.write(doc_path, xml_declaration=True, encoding="UTF-8")

        # 6. Rezipper
        with zipfile.ZipFile(dst_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for dp, _, fns in os.walk(tmp):
                for fn in fns:
                    f = os.path.join(dp, fn)
                    zout.write(f, os.path.relpath(f, tmp))

    return True


class ApplyDocxCropsTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage, None, None]:
        files = tool_parameters.get("files", [])
        if not files:
            yield self.create_text_message("No files provided")
            yield self.create_json_message({"status": "error", "message": "No files provided", "results": []})
            return

        json_results = []

        for file in files:
            file_extension = (file.extension or ".tmp").lower()

            if file_extension not in (".docx", ".docm"):
                yield self.create_text_message(f"Skipped {file.filename}: not a .docx/.docm file")
                json_results.append({"filename": file.filename, "status": "skipped", "reason": "not a docx/docm"})
                continue

            src_path = None
            dst_path = None
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=file_extension) as src_f:
                    src_f.write(file.blob)
                    src_path = src_f.name

                dst_path = src_path + "_cropped" + file_extension
                crop_applied = apply_docx_crops(src_path, dst_path)

                if crop_applied:
                    with open(dst_path, "rb") as out_f:
                        cropped_bytes = out_f.read()
                    yield self.create_blob_message(
                        cropped_bytes,
                        meta={"mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
                    )
                    json_results.append({"filename": file.filename, "status": "success", "crops_applied": True})
                else:
                    # Aucun crop détecté : retourner le fichier original tel quel
                    yield self.create_blob_message(
                        file.blob,
                        meta={"mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
                    )
                    json_results.append({"filename": file.filename, "status": "success", "crops_applied": False})

            except Exception as e:
                error_msg = f"Error processing {file.filename}: {str(e)}"
                yield self.create_text_message(error_msg)
                json_results.append({"filename": file.filename, "status": "error", "error": error_msg})

            finally:
                for path in [src_path, dst_path]:
                    if path and os.path.exists(path):
                        os.unlink(path)

        yield self.create_json_message({
            "status": "success",
            "total_files": len(files),
            "results": json_results,
        })