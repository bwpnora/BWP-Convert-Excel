"""
BWPConvertTTNVN Direct OpenXML Ribbon Packaging Pipeline
Injects customUI/customUI14.xml into .xlam OpenXML ZIP packages directly
without relying on Excel COM to manipulate the ribbon.
"""

import argparse
import hashlib
import os
import shutil
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Optional, Union

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_RIBBON_XML = PROJECT_ROOT / "ribbon" / "customUI14.xml"
DEFAULT_XLAM_PATH = PROJECT_ROOT / "dist" / "BWPConvertTTNVN.xlam"

RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
CUSTOMUI_REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
CUSTOMUI_TARGET = "customUI/customUI14.xml"


def inject_ribbon(
    xlam_path: Union[str, Path],
    ribbon_xml_path: Optional[Union[str, Path]] = None,
) -> Path:
    """
    Directly injects customUI14.xml into an .xlam OpenXML ZIP package.

    Performs:
    1. Reads and validates ribbon XML (UTF-8, strips BOM if present).
    2. Opens .xlam as ZIP, hashes xl/vbaProject.bin if present to guarantee integrity.
    3. Adds or updates relationship in _rels/.rels targeting customUI/customUI14.xml with unique Id.
    4. Ensures [Content_Types].xml recognizes .xml parts.
    5. Injects customUI/customUI14.xml part.
    6. Verifies xl/vbaProject.bin preservation and package integrity.

    Returns the Path of the modified .xlam file.
    """
    target_xlam = Path(xlam_path).resolve()
    if not target_xlam.is_file():
        raise FileNotFoundError(f"Target add-in file does not exist: {target_xlam}")

    target_ribbon = Path(ribbon_xml_path).resolve() if ribbon_xml_path else DEFAULT_RIBBON_XML
    if not target_ribbon.is_file():
        raise FileNotFoundError(f"Ribbon XML definition does not exist: {target_ribbon}")

    # 1. Read ribbon XML content
    ribbon_bytes = target_ribbon.read_bytes()
    if ribbon_bytes.startswith(b"\xef\xbb\xbf"):
        ribbon_bytes = ribbon_bytes[3:]  # Strip UTF-8 BOM if present

    # Validate well-formedness of ribbon XML
    try:
        ET.fromstring(ribbon_bytes)
    except ET.ParseError as e:
        raise ValueError(f"Invalid ribbon XML in {target_ribbon}: {e}")

    temp_packaged = target_xlam.with_name(f"{target_xlam.stem}.tmp_ribbon{target_xlam.suffix}")
    vba_sha256_before: Optional[str] = None

    ET.register_namespace("", RELS_NS)
    ET.register_namespace("", CONTENT_TYPES_NS)

    with zipfile.ZipFile(target_xlam, "r") as zin:
        namelist = zin.namelist()

        # Check for vbaProject.bin and record hash
        if "xl/vbaProject.bin" in namelist:
            vba_bytes_before = zin.read("xl/vbaProject.bin")
            vba_sha256_before = hashlib.sha256(vba_bytes_before).hexdigest()

        with zipfile.ZipFile(temp_packaged, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == CUSTOMUI_TARGET:
                    continue
                content = zin.read(item.filename)

                if item.filename == "_rels/.rels":
                    rels_root = ET.fromstring(content)
                    existing_customui_rel = None
                    existing_ids = set()

                    for rel in rels_root:
                        r_id = rel.attrib.get("Id", "")
                        if r_id:
                            existing_ids.add(r_id)
                        if (
                            rel.attrib.get("Target") == CUSTOMUI_TARGET
                            or rel.attrib.get("Type") == CUSTOMUI_REL_TYPE
                        ):
                            existing_customui_rel = rel

                    if existing_customui_rel is not None:
                        existing_customui_rel.set("Type", CUSTOMUI_REL_TYPE)
                        existing_customui_rel.set("Target", CUSTOMUI_TARGET)
                    else:
                        rel_id = "rIdCustomUI"
                        suffix = 1
                        while rel_id in existing_ids:
                            rel_id = f"rIdCustomUI{suffix}"
                            suffix += 1

                        new_rel = ET.SubElement(rels_root, f"{{{RELS_NS}}}Relationship")
                        new_rel.set("Id", rel_id)
                        new_rel.set("Type", CUSTOMUI_REL_TYPE)
                        new_rel.set("Target", CUSTOMUI_TARGET)

                    content = (
                        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                        + ET.tostring(rels_root, encoding="utf-8")
                    )

                elif item.filename == "[Content_Types].xml":
                    ct_root = ET.fromstring(content)
                    has_xml_type = False
                    for el in ct_root:
                        if el.attrib.get("Extension", "").lower() == "xml":
                            has_xml_type = True
                            break
                        if el.attrib.get("PartName") == f"/{CUSTOMUI_TARGET}":
                            has_xml_type = True
                            break

                    if not has_xml_type:
                        new_default = ET.SubElement(ct_root, f"{{{CONTENT_TYPES_NS}}}Default")
                        new_default.set("Extension", "xml")
                        new_default.set("ContentType", "application/xml")

                    content = (
                        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                        + ET.tostring(ct_root, encoding="utf-8")
                    )

                zout.writestr(item, content)

            # Injects customUI/customUI14.xml part
            zout.writestr(CUSTOMUI_TARGET, ribbon_bytes)

    # Replace original xlam with newly packaged xlam
    shutil.move(temp_packaged, target_xlam)

    # Post-packaging verification
    with zipfile.ZipFile(target_xlam, "r") as zver:
        ver_names = zver.namelist()

        if CUSTOMUI_TARGET not in ver_names:
            raise RuntimeError(f"Packaging failed: {CUSTOMUI_TARGET} missing from final package.")

        injected_content = zver.read(CUSTOMUI_TARGET)
        if injected_content != ribbon_bytes:
            raise RuntimeError("Packaging failed: injected ribbon XML does not match source bytes.")

        if vba_sha256_before is not None:
            if "xl/vbaProject.bin" not in ver_names:
                raise RuntimeError("Packaging corrupted archive: xl/vbaProject.bin was lost.")
            vba_sha256_after = hashlib.sha256(zver.read("xl/vbaProject.bin")).hexdigest()
            if vba_sha256_before != vba_sha256_after:
                raise RuntimeError(
                    f"VBA project integrity mismatch! "
                    f"Before: {vba_sha256_before}, After: {vba_sha256_after}"
                )

        rels_bytes = zver.read("_rels/.rels")
        rels_root = ET.fromstring(rels_bytes)
        has_customui_rel = any(
            rel.attrib.get("Target") == CUSTOMUI_TARGET
            and rel.attrib.get("Type") == CUSTOMUI_REL_TYPE
            for rel in rels_root
        )
        if not has_customui_rel:
            raise RuntimeError("Packaging failed: customUI relationship missing from _rels/.rels")

    return target_xlam


def main():
    parser = argparse.ArgumentParser(
        description="Direct OpenXML Ribbon Packaging Pipeline for BWPConvertTTNVN"
    )
    parser.add_argument(
        "xlam",
        nargs="?",
        default=str(DEFAULT_XLAM_PATH),
        help=f"Path to .xlam file (default: {DEFAULT_XLAM_PATH})",
    )
    parser.add_argument(
        "--ribbon-xml",
        default=str(DEFAULT_RIBBON_XML),
        help=f"Path to customUI XML definition (default: {DEFAULT_RIBBON_XML})",
    )

    args = parser.parse_args()

    xlam_path = Path(args.xlam)
    ribbon_path = Path(args.ribbon_xml)

    try:
        out = inject_ribbon(xlam_path, ribbon_path)
        print(f"[SUCCESS] Ribbon XML successfully injected into: {out}")
    except Exception as e:
        print(f"[ERROR] Failed to inject ribbon: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
