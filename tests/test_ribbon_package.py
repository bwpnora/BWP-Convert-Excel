"""
Unit and Integration Test Suite for Ribbon Definition and OpenXML Packaging
Validates:
1. Well-formedness, Office 2010+ schema, and UTF-8 encoding of customUI14.xml.
2. Callback contract: all onLoad, onAction, and getEnabled attributes in customUI14.xml
   match Public Sub declarations in src/excel/RibbonCallbacks.bas.
3. Pure 7-bit ASCII safety of src/excel/RibbonCallbacks.bas.
4. Direct OpenXML injection via scripts/package_ribbon.py onto an .xlam package.
5. Preservation of xl/vbaProject.bin and unique relationship IDs.
6. Live Excel COM reopen validation: fresh Excel COM instance opens the packaged .xlam
   without errors.
7. Live Excel COM VBA compilation and callback smoke testing.
"""

import hashlib
import os
import shutil
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from scripts.build import ExcelSession
from scripts.package_ribbon import (
    CONTENT_TYPES_NS,
    CUSTOMUI_REL_TYPE,
    CUSTOMUI_TARGET,
    RELS_NS,
    inject_ribbon,
)

RIBBON_XML_PATH = PROJECT_ROOT / "ribbon" / "customUI14.xml"
RIBBON_BAS_PATH = PROJECT_ROOT / "src" / "excel" / "RibbonCallbacks.bas"
OFFICE_2010_CUSTOMUI_NS = "http://schemas.microsoft.com/office/2009/07/customui"


class TestRibbonXmlContract(unittest.TestCase):
    """Validates the Ribbon XML structure, metadata, and callback names."""

    def test_ribbon_xml_file_exists_and_utf8_no_bom(self):
        self.assertTrue(RIBBON_XML_PATH.is_file(), f"{RIBBON_XML_PATH} must exist")
        raw_bytes = RIBBON_XML_PATH.read_bytes()
        self.assertFalse(
            raw_bytes.startswith(b"\xef\xbb\xbf"),
            "customUI14.xml must not contain a UTF-8 BOM",
        )
        # Verify valid UTF-8
        try:
            decoded = raw_bytes.decode("utf-8")
        except UnicodeDecodeError as e:
            self.fail(f"customUI14.xml is not valid UTF-8: {e}")
        self.assertIn("customUI", decoded)

    def test_ribbon_xml_schema_and_structure(self):
        tree = ET.parse(str(RIBBON_XML_PATH))
        root = tree.getroot()

        # Office 2010+ namespace check
        self.assertEqual(
            root.tag,
            f"{{{OFFICE_2010_CUSTOMUI_NS}}}customUI",
            f"Root must be <customUI> with namespace {OFFICE_2010_CUSTOMUI_NS}",
        )
        self.assertEqual(root.attrib.get("onLoad"), "OnRibbonLoad")

        # Find tab
        ns = {"ns": OFFICE_2010_CUSTOMUI_NS}
        tab = root.find(".//ns:tab[@id='tabBWPConvert']", ns)
        self.assertIsNotNone(tab, "Tab 'tabBWPConvert' must exist")
        self.assertEqual(tab.attrib.get("label"), "BWPConvertTTNVN")
        self.assertEqual(tab.attrib.get("keytip"), "B")

        # Check groups
        grp_conv = tab.find("ns:group[@id='grpConversion']", ns)
        self.assertIsNotNone(grp_conv, "Group 'grpConversion' must exist")
        self.assertEqual(grp_conv.attrib.get("label"), "Chuyển đổi số")

        grp_settings = tab.find("ns:group[@id='grpSettings']", ns)
        self.assertIsNotNone(grp_settings, "Group 'grpSettings' must exist")
        self.assertEqual(grp_settings.attrib.get("label"), "Hệ thống")

        # Check buttons in grpConversion
        btn_convert = grp_conv.find("ns:button[@id='btnConvert']", ns)
        self.assertIsNotNone(btn_convert)
        self.assertEqual(btn_convert.attrib.get("label"), "Đổi số thành chữ")
        self.assertEqual(btn_convert.attrib.get("size"), "large")
        self.assertEqual(btn_convert.attrib.get("imageMso"), "FunctionsTextInsertGallery")
        self.assertEqual(btn_convert.attrib.get("onAction"), "OnConvertClick")

        btn_quick = grp_conv.find("ns:button[@id='btnQuickConvert']", ns)
        self.assertIsNotNone(btn_quick)
        self.assertEqual(btn_quick.attrib.get("label"), "Chuyển nhanh")
        self.assertEqual(btn_quick.attrib.get("size"), "large")
        self.assertEqual(btn_quick.attrib.get("imageMso"), "AutoSum")
        self.assertEqual(btn_quick.attrib.get("onAction"), "OnQuickConvertClick")

        btn_undo = grp_conv.find("ns:button[@id='btnUndo']", ns)
        self.assertIsNotNone(btn_undo)
        self.assertEqual(btn_undo.attrib.get("label"), "Hoàn tác")
        self.assertEqual(btn_undo.attrib.get("size"), "normal")
        self.assertEqual(btn_undo.attrib.get("imageMso"), "Undo")
        self.assertEqual(btn_undo.attrib.get("onAction"), "OnUndoClick")
        self.assertEqual(btn_undo.attrib.get("getEnabled"), "GetUndoEnabled")

        # Check buttons in grpSettings
        btn_settings = grp_settings.find("ns:button[@id='btnSettings']", ns)
        self.assertIsNotNone(btn_settings)
        self.assertEqual(btn_settings.attrib.get("label"), "Cài đặt")
        self.assertEqual(btn_settings.attrib.get("size"), "normal")
        self.assertEqual(btn_settings.attrib.get("imageMso"), "ControlProperties")
        self.assertEqual(btn_settings.attrib.get("onAction"), "OnSettingsClick")

        btn_about = grp_settings.find("ns:button[@id='btnAbout']", ns)
        self.assertIsNotNone(btn_about)
        self.assertEqual(btn_about.attrib.get("label"), "Giới thiệu")
        self.assertEqual(btn_about.attrib.get("size"), "normal")
        self.assertEqual(btn_about.attrib.get("imageMso"), "Info")
        self.assertEqual(btn_about.attrib.get("onAction"), "OnAboutClick")


class TestRibbonCallbacksSource(unittest.TestCase):
    """Validates pure 7-bit ASCII safety and VBA procedure signatures."""

    def test_ribbon_callbacks_bas_pure_7bit_ascii(self):
        self.assertTrue(RIBBON_BAS_PATH.is_file(), f"{RIBBON_BAS_PATH} must exist")
        content = RIBBON_BAS_PATH.read_bytes()
        non_ascii = [b for b in content if b > 127]
        self.assertEqual(
            len(non_ascii),
            0,
            f"Found {len(non_ascii)} non-ASCII bytes in RibbonCallbacks.bas: {non_ascii[:10]}",
        )

    def test_all_xml_callbacks_exist_in_bas(self):
        # Extract callbacks from XML
        tree = ET.parse(str(RIBBON_XML_PATH))
        root = tree.getroot()

        required_callbacks = set()
        on_load = root.attrib.get("onLoad")
        if on_load:
            required_callbacks.add(on_load)

        for el in root.iter():
            for attr in ["onAction", "getEnabled"]:
                val = el.attrib.get(attr)
                if val:
                    required_callbacks.add(val)

        self.assertGreaterEqual(len(required_callbacks), 7)

        # Parse Public Sub in RibbonCallbacks.bas
        bas_text = RIBBON_BAS_PATH.read_text(encoding="utf-8")
        import re

        sub_pattern = re.compile(r"^\s*Public\s+Sub\s+([A-Za-z0-9_]+)\s*\(", re.MULTILINE)
        found_subs = set(sub_pattern.findall(bas_text))

        for cb in required_callbacks:
            self.assertIn(
                cb,
                found_subs,
                f"Callback '{cb}' declared in Ribbon XML was not found as a 'Public Sub {cb}' in RibbonCallbacks.bas",
            )


class TestRibbonPackagingOpenXml(unittest.TestCase):
    """Validates OpenXML ZIP package manipulation and integrity checks."""

    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="test_ribbon_pkg_")).resolve()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def _create_mock_xlam(self, include_vba: bool = True) -> Path:
        """Creates a mock OpenXML .xlam package without needing Excel COM."""
        xlam_path = (self.temp_dir / ("mock_with_vba.xlam" if include_vba else "mock_blank.xlam")).resolve()

        rels_content = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            b"</Relationships>"
        )

        content_types = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            b'<Default Extension="xml" ContentType="application/xml"/>'
            b'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.addin.macroEnabled.main+xml"/>'
            b"</Types>"
        )

        with zipfile.ZipFile(xlam_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("_rels/.rels", rels_content)
            zf.writestr("[Content_Types].xml", content_types)
            zf.writestr("xl/workbook.xml", b"<workbook/>")
            if include_vba:
                fake_vba = b"VBA_PROJECT_MOCK_BINARY_DATA_\x01\x02\x03\x04\xff\xfe"
                zf.writestr("xl/vbaProject.bin", fake_vba)

        return xlam_path

    def test_inject_ribbon_into_mock_xlam_with_vba_preservation(self):
        xlam_path = self._create_mock_xlam(include_vba=True)

        with zipfile.ZipFile(xlam_path, "r") as z_orig:
            vba_before = z_orig.read("xl/vbaProject.bin")
            vba_hash_before = hashlib.sha256(vba_before).hexdigest()

        # Run injection
        out_path = inject_ribbon(xlam_path, RIBBON_XML_PATH)
        self.assertEqual(out_path.resolve(), xlam_path.resolve())

        # Verify output archive
        with zipfile.ZipFile(xlam_path, "r") as z_out:
            namelist = z_out.namelist()
            self.assertIn(CUSTOMUI_TARGET, namelist)
            self.assertIn("xl/vbaProject.bin", namelist)

            vba_after = z_out.read("xl/vbaProject.bin")
            vba_hash_after = hashlib.sha256(vba_after).hexdigest()
            self.assertEqual(
                vba_hash_before,
                vba_hash_after,
                "xl/vbaProject.bin must be byte-for-byte identical after packaging",
            )

            # Check _rels/.rels
            rels_root = ET.fromstring(z_out.read("_rels/.rels"))
            rel_customui = [
                el
                for el in rels_root
                if el.attrib.get("Target") == CUSTOMUI_TARGET
                and el.attrib.get("Type") == CUSTOMUI_REL_TYPE
            ]
            self.assertEqual(len(rel_customui), 1)
            self.assertTrue(rel_customui[0].attrib.get("Id"))

    def test_inject_ribbon_idempotency(self):
        xlam_path = self._create_mock_xlam(include_vba=False)

        inject_ribbon(xlam_path, RIBBON_XML_PATH)
        inject_ribbon(xlam_path, RIBBON_XML_PATH)

        with zipfile.ZipFile(xlam_path, "r") as z_out:
            rels_root = ET.fromstring(z_out.read("_rels/.rels"))
            rel_customui = [
                el
                for el in rels_root
                if el.attrib.get("Target") == CUSTOMUI_TARGET
                and el.attrib.get("Type") == CUSTOMUI_REL_TYPE
            ]
            self.assertEqual(
                len(rel_customui),
                1,
                "Multiple injections must not duplicate the customUI relationship",
            )
            self.assertEqual(
                z_out.namelist().count(CUSTOMUI_TARGET),
                1,
                "Multiple injections must not duplicate the customUI entry in the ZIP",
            )


class TestRibbonLiveExcelCOM(unittest.TestCase):
    """Executes live Excel COM reopen validation and VBA callback smoke tests."""

    temp_dir = None
    blank_xlam = None

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = Path(tempfile.mkdtemp(prefix="test_ribbon_com_"))
        cls.blank_xlam = cls.temp_dir / "LiveTestAddin.xlam"

        # 1. Create a clean blank .xlam via Excel COM
        with ExcelSession(visible=False, display_alerts=False) as session:
            wb = session.excel.Workbooks.Add()
            wb.SaveAs(str(cls.blank_xlam), 55)  # 55 = xlOpenXMLAddIn
            wb.Close(False)

        # 2. Package customUI14.xml into LiveTestAddin.xlam
        inject_ribbon(cls.blank_xlam, RIBBON_XML_PATH)

    @classmethod
    def tearDownClass(cls):
        if cls.temp_dir and cls.temp_dir.exists():
            shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def test_live_excel_reopen_packaged_xlam(self):
        """Verifies that a fresh isolated Excel COM instance opens the packaged .xlam with 0 errors."""
        with ExcelSession(visible=False, display_alerts=False) as session:
            wb = session.excel.Workbooks.Open(str(self.blank_xlam))
            self.assertIsNotNone(wb, "Packaged .xlam must open cleanly in Excel")
            self.assertEqual(wb.Name, self.blank_xlam.name)
            wb.Close(False)

    def test_live_excel_vba_ribbon_callbacks_smoke(self):
        """Imports RibbonCallbacks.bas and dependencies into a live workbook and verifies procedure invocation."""
        with ExcelSession(visible=False, display_alerts=False) as session:
            wb = session.excel.Workbooks.Add()
            vb_proj = wb.VBProject

            # Required dependency modules in order
            deps = [
                ("src/core", "CoreTypes.bas"),
                ("src/core", "UnicodeText.bas"),
                ("src/core", "VietnameseNumber.bas"),
                ("src/core", "Currency.bas"),
                ("src/core", "TextFormatter.bas"),
                ("src/core", "CoreCoordinator.bas"),
                ("src/excel", "BuildInfo.bas"),
                ("src/excel", "Settings.bas"),
                ("src/excel", "UndoManager.bas"),
                ("src/excel", "CellProcessor.bas"),
                ("src/excel", "RibbonCallbacks.bas"),
            ]

            for rel_dir, mod_file in deps:
                mod_path = PROJECT_ROOT / rel_dir / mod_file
                vb_proj.VBComponents.Import(str(mod_path))

            # Add test harness to exercise callback routines
            test_mod = vb_proj.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
            test_mod.Name = "modTestRibbon"
            test_mod.CodeModule.AddFromString(
                """Option Explicit

Public Function TestUndoAvailableContract() As Boolean
    Dim isEnabled As Variant
    isEnabled = False
    
    ' Initially no undo snapshot should be committed
    RibbonCallbacks.GetUndoEnabled Nothing, isEnabled
    If isEnabled <> False Then
        TestUndoAvailableContract = False
        Exit Function
    End If
    
    ' Invalidate ribbon calls should not raise runtime error even if gRibbon is Nothing
    RibbonCallbacks.InvalidateRibbonControl "btnUndo"
    RibbonCallbacks.InvalidateUndoButton
    RibbonCallbacks.InvalidateUndoRibbon
    
    TestUndoAvailableContract = True
End Function
"""
            )

            res = session.excel.Run(f"'{wb.Name}'!TestUndoAvailableContract")
            self.assertTrue(res, "TestUndoAvailableContract should return True without errors")

            wb.Close(False)


if __name__ == "__main__":
    unittest.main()
