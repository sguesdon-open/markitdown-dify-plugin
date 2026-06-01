from collections.abc import Generator
from typing import Any
import tempfile
import os
import zipfile
import xml.etree.ElementTree as ET
import io

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
    """Word stocke les valeurs srcRect en 1/100 000 (ex: 10000 = 10%).
    Les valeurs négatives (zoom-out Word) sont clampées à 0."""
    return max(0.0, int(val) / 100_000.0)


def _has_real_crop(attrs: dict) -> bool:
    """Retourne True si au moins une valeur srcRect est positive (crop réel).
    Les srcRect avec uniquement des valeurs nulles ou négatives sont ignorés."""
    return any(int(v) > 0 for v in attrs.values())


def apply_docx_crops(src_path: str, dst_path: str) -> bool:
    """
    Lit src_path (.docx), applique physiquement les crops Word (a:srcRect)
    sur chaque image avec Pillow, supprime le srcRect du XML, et écrit dst_path.
    Retourne True si au moins une image a été croppée.

    Corrections vs version précédente :
    - Images partagées avec crops différents (sprite sheets) : crée une nouvelle
      image par usage et met à jour les relations, au lieu d'écraser le fichier
      source (ce qui cassait les usages suivants de la même image).
    - Valeurs srcRect négatives (zoom-out Word) : clampées à 0, srcRect ignoré
      si aucune valeur positive.
    """
    try:
        from PIL import Image
    except ImportError:
        return False

    from pathlib import Path

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)

        # 1. Extraire le docx
        with zipfile.ZipFile(src_path, "r") as z:
            z.extractall(tmp)

        # 2. Lire les relations
        rels_path = tmp / "word" / "_rels" / "document.xml.rels"
        if not rels_path.exists():
            return False

        rels_tree = ET.parse(rels_path)
        rels_root = rels_tree.getroot()
        rel_map = {r.get("Id"): r.get("Target") for r in rels_root}

        # Calculer le prochain rId disponible (pour les nouvelles relations)
        existing_ids = [
            int(r.get("Id", "rId0").replace("rId", ""))
            for r in rels_root if (r.get("Id") or "").startswith("rId")
        ]
        _next_rid = [max(existing_ids, default=0) + 1]

        def _new_rid():
            rid = f"rId{_next_rid[0]}"
            _next_rid[0] += 1
            return rid

        # 3. Parser document.xml
        doc_path = tmp / "word" / "document.xml"
        doc_tree = ET.parse(doc_path)
        doc_root = doc_tree.getroot()

        # 4. Parcourir les blipFill avec srcRect
        crops_applied = 0

        for blipFill in doc_root.iter(f"{{{_NS['pic']}}}blipFill"):
            src_rect = blipFill.find(f"{{{_NS['a']}}}srcRect")
            if src_rect is None:
                continue

            # Ne récupérer que les attributs présents dans le XML
            attrs = {k: src_rect.get(k, "0")
                     for k in ("l", "t", "r", "b") if src_rect.get(k) is not None}

            # Ignorer si pas de crop réel (vide ou que des valeurs négatives/nulles)
            if not attrs or not _has_real_crop(attrs):
                continue

            blip = blipFill.find(f"{{{_NS['a']}}}blip")
            if blip is None:
                continue

            embed = blip.get(f"{{{_NS['r']}}}embed")
            if embed not in rel_map:
                continue

            img_rel = rel_map[embed]
            img_abs = tmp / "word" / img_rel
            if not img_abs.exists():
                continue

            try:
                img = Image.open(img_abs)
                fmt = img.format or "PNG"
                w, h = img.size

                l = _parse_crop_pct(attrs.get("l", "0"))
                t = _parse_crop_pct(attrs.get("t", "0"))
                r = _parse_crop_pct(attrs.get("r", "0"))
                b = _parse_crop_pct(attrs.get("b", "0"))

                # Clamp pour garantir des coordonnées valides
                left   = max(0, int(w * l))
                top    = max(0, int(h * t))
                right  = min(w, int(w * (1.0 - r)))
                bottom = min(h, int(h * (1.0 - b)))

                if left >= right or top >= bottom:
                    continue

                cropped = img.crop((left, top, right, bottom))
                save_fmt = fmt if fmt in ("PNG", "JPEG", "GIF", "BMP", "TIFF") else "PNG"

                # Créer une NOUVELLE image pour ce crop (ne pas écraser l'originale,
                # qui peut être réutilisée avec un srcRect différent ailleurs)
                new_name = f"image_crop_{crops_applied + 1}{img_abs.suffix}"
                new_img_path = tmp / "word" / "media" / new_name

                out = io.BytesIO()
                cropped.save(out, format=save_fmt)
                new_img_path.write_bytes(out.getvalue())

                # Nouvelle relation pointant vers la nouvelle image
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

            except Exception:
                continue  # on garde l'image originale si erreur

        if not crops_applied:
            return False

        # 5. Réécrire document.xml ET les relations
        doc_tree.write(doc_path, xml_declaration=True, encoding="UTF-8")
        rels_tree.write(rels_path, xml_declaration=True, encoding="UTF-8")

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