"""
BWPConvertTTNVN Automated 3-Tier Test Orchestrator
Executes full verification against the compiled dist/BWPConvertTTNVN.xlam add-in:
- Tier 1: Core API Tests (ConvertNumberDefault, ConvertNumberWithOptions across 66 expected_cases.csv)
- Tier 2: Real Worksheet UDF Tests (=BWPVND, =BWPVNWORDS, =BWPVNDUPPER, aliases, blanks, errors)
- Tier 3: Excel Integration Tests (OnQuickConvertClick, destination expansion, cross-sheet, overlap rejection, Undo, skip preservation)

Outputs a comprehensive test report to console and writes dist/test-report.txt.
"""

import argparse
import csv
import hashlib
import os
import sys
import time
import unittest
import winreg
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from scripts.build import DEFAULT_XLAM_PATH, ExcelSession

DEFAULT_CSV_PATH = PROJECT_ROOT / "tests" / "expected_cases.csv"
DEFAULT_REPORT_PATH = PROJECT_ROOT / "dist" / "test-report.txt"

# Target XLAM path configurable via environment variable or CLI argument
CURRENT_XLAM_PATH = Path(os.environ.get("BWP_TEST_XLAM", str(DEFAULT_XLAM_PATH)))

XL_ERR_VALUE = -2146826273  # Excel xlErrValue (#VALUE!)
XL_ERR_VALUE_CODES = (-2146826273, -2146826281)


# ---------------------------------------------------------------------------
# Test Results Collector
# ---------------------------------------------------------------------------

class TestRecord:
    def __init__(self, tier: str, name: str, passed: bool, detail: str = "", duration: float = 0.0):
        self.tier = tier
        self.name = name
        self.passed = passed
        self.detail = detail
        self.duration = duration


GLOBAL_TEST_RECORDS: List[TestRecord] = []


def record_result(tier: str, name: str, passed: bool, detail: str = "", duration: float = 0.0):
    GLOBAL_TEST_RECORDS.append(TestRecord(tier, name, passed, detail, duration))


# ---------------------------------------------------------------------------
# Tier 1: Core API Tests
# ---------------------------------------------------------------------------

class Tier1CoreApiTests(unittest.TestCase):
    """
    Tier 1: Core API Tests operating on compiled BWPConvertTTNVN.xlam via fully qualified macro names:
    - 'BWPConvertTTNVN.xlam'!ConvertNumberDefault
    - 'BWPConvertTTNVN.xlam'!ConvertNumberWithOptions
    - All 66+ test scenarios in tests/expected_cases.csv
    """

    session: Optional[ExcelSession] = None
    addin_wb: Any = None
    addin_name: str = ""

    @classmethod
    def setUpClass(cls):
        target = Path(CURRENT_XLAM_PATH).resolve()
        if not target.is_file():
            raise FileNotFoundError(f"Target add-in not found: {target}. Run scripts/build.py first.")
        cls.session = ExcelSession(visible=False, display_alerts=False)
        cls.session.__enter__()
        cls.addin_wb = cls.session.excel.Workbooks.Open(str(target))
        cls.addin_name = cls.addin_wb.Name

    @classmethod
    def tearDownClass(cls):
        if cls.addin_wb:
            try:
                cls.addin_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if cls.session:
            cls.session.__exit__(None, None, None)

    def test_convert_number_default_smoke(self):
        """Verify 'BWPConvertTTNVN.xlam'!ConvertNumberDefault on key values."""
        t0 = time.time()
        macro = f"'{self.addin_name}'!ConvertNumberDefault"

        cases = [
            (125430000, "Một trăm hai mươi lăm triệu bốn trăm ba mươi nghìn đồng chẵn."),
            (0, "Không đồng chẵn."),
            (1000, "Một nghìn đồng chẵn."),
            (500000, "Năm trăm nghìn đồng chẵn."),
            (-100, "Âm một trăm đồng chẵn."),
            (1000000000, "Một tỷ đồng chẵn."),
            (1000000000000, "Một nghìn tỷ đồng chẵn."),
        ]

        for val, expected in cases:
            with self.subTest(val=val):
                res = self.session.excel.Run(macro, val)
                self.assertEqual(res, expected, f"Failed for val={val}")

        record_result("Tier 1", "ConvertNumberDefault smoke test (7 cases)", True, duration=time.time() - t0)

    def test_expected_cases_matrix(self):
        """Execute all rows in tests/expected_cases.csv via ConvertNumberWithOptions."""
        t0 = time.time()
        csv_file = DEFAULT_CSV_PATH
        self.assertTrue(csv_file.is_file(), f"expected_cases.csv not found: {csv_file}")

        with open(csv_file, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            rows = list(reader)

        self.assertGreaterEqual(len(rows), 60, f"Expected >= 60 test cases, found {len(rows)}")

        macro = f"'{self.addin_name}'!ConvertNumberWithOptions"
        passed_count = 0

        for row in rows:
            case_id = row["case_id"]
            raw_input = row["input"]
            expected = row["expected"]
            zero_style = int(row["zero_style"])
            thousand_style = int(row["thousand_style"])
            four_style = int(row["four_style"])
            dec_mode = int(row["decimal_mode"])
            expected_error = row["expected_error"].strip()

            if expected_error == "ERR_INVALID_NUMBER":
                input_val = raw_input
            else:
                input_val = float(raw_input) if "." in raw_input else int(raw_input)

            with self.subTest(case_id=case_id, input=raw_input):
                # CurrType=2 (VnCurrNone), AddChan=False, DecPlaces=0, Casing=2 (VnCaseLower), AddPeriod=False
                res = self.session.excel.Run(
                    macro,
                    input_val,
                    2,      # VnCurrNone
                    False,  # AddChan
                    0,      # DecPlaces
                    zero_style,
                    thousand_style,
                    four_style,
                    dec_mode,
                    2,      # VnCaseLower
                    False,  # AddPeriod
                )

                if expected_error:
                    self.assertEqual(res, "", f"Expected empty string for error {expected_error}, got: '{res}'")
                else:
                    self.assertEqual(res, expected, f"Mismatch in {case_id} for input {raw_input}")
                passed_count += 1

        record_result(
            "Tier 1",
            f"ConvertNumberWithOptions across expected_cases.csv ({passed_count} scenarios)",
            True,
            duration=time.time() - t0,
        )

    def test_dialects_and_casing_options(self):
        """Verify dialect styles and casing styles via ConvertNumberWithOptions."""
        t0 = time.time()
        macro = f"'{self.addin_name}'!ConvertNumberWithOptions"

        # Linh vs Lẻ
        res_le = self.session.excel.Run(macro, 105, 0, True, 0, 0, 0, 0, 0, 0, True)
        res_linh = self.session.excel.Run(macro, 105, 0, True, 0, 1, 0, 0, 0, 0, True)
        self.assertIn("lẻ", res_le.lower())
        self.assertIn("linh", res_linh.lower())

        # Ngàn vs Nghìn
        res_nghin = self.session.excel.Run(macro, 1000, 0, True, 0, 0, 0, 0, 0, 0, True)
        res_ngan = self.session.excel.Run(macro, 1000, 0, True, 0, 0, 1, 0, 0, 0, True)
        self.assertIn("nghìn", res_nghin.lower())
        self.assertIn("ngàn", res_ngan.lower())

        # Casing: Sentence (0), Upper (1), Lower (2)
        res_upper = self.session.excel.Run(macro, 1000, 0, True, 0, 0, 0, 0, 0, 1, True)
        self.assertEqual(res_upper, "MỘT NGHÌN ĐỒNG CHẴN.")

        record_result("Tier 1", "Dialect and casing styles verification", True, duration=time.time() - t0)


# ---------------------------------------------------------------------------
# Tier 2: Real Worksheet UDF Tests
# ---------------------------------------------------------------------------

class Tier2WorksheetUdfTests(unittest.TestCase):
    """
    Tier 2: Real Worksheet UDF Tests operating with the add-in open:
    - Canonical live formulas: =BWPVND(A1), =BWPVNWORDS(A1), =BWPVNDUPPER(A1)
    - Convenience aliases: =VND(A1), =VNWORDS(A1), =VNDUPPER(A1)
    - CalculateFull() verification
    - Decimal reading in =BWPVNWORDS(125.05)
    - Blank cell returns ""
    - Overflow (>999T) and non-numeric returns #VALUE!
    """

    session: Optional[ExcelSession] = None
    addin_wb: Any = None
    test_wb: Any = None
    ws: Any = None

    @classmethod
    def setUpClass(cls):
        target = Path(CURRENT_XLAM_PATH).resolve()
        if not target.is_file():
            raise FileNotFoundError(f"Target add-in not found: {target}")
        cls.session = ExcelSession(visible=False, display_alerts=False)
        cls.session.__enter__()
        cls.addin_wb = cls.session.excel.Workbooks.Open(str(target))
        cls.test_wb = cls.session.excel.Workbooks.Add()
        cls.ws = cls.test_wb.Worksheets(1)

    @classmethod
    def tearDownClass(cls):
        if cls.test_wb:
            try:
                cls.test_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if cls.addin_wb:
            try:
                cls.addin_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if cls.session:
            cls.session.__exit__(None, None, None)

    def setUp(self):
        self.ws.Cells.Clear()

    def test_canonical_bwpvnd_formula(self):
        """=BWPVND(A1) produces sentence-case Vietnamese Dong currency phrase."""
        t0 = time.time()
        self.ws.Range("A1").Value = 125430000
        self.ws.Range("B1").Formula = "=BWPVND(A1)"
        self.session.excel.CalculateFull()

        expected = "Một trăm hai mươi lăm triệu bốn trăm ba mươi nghìn đồng chẵn."
        self.assertEqual(self.ws.Range("B1").Value2, expected)
        record_result("Tier 2", "Canonical =BWPVND(A1) formula", True, duration=time.time() - t0)

    def test_canonical_bwpvnwords_decimal_formula(self):
        """=BWPVNWORDS(A2) reads decimal digits with 'phẩy'."""
        t0 = time.time()
        self.ws.Range("A2").Value = 125.05
        self.ws.Range("B2").Formula = "=BWPVNWORDS(A2)"
        self.session.excel.CalculateFull()

        expected = "Một trăm hai mươi lăm phẩy không năm."
        self.assertEqual(self.ws.Range("B2").Value2, expected)
        record_result("Tier 2", "Canonical =BWPVNWORDS(A2) decimal reading", True, duration=time.time() - t0)

    def test_canonical_bwpvndupper_formula(self):
        """=BWPVNDUPPER(A3) produces all-uppercase Vietnamese Dong phrase."""
        t0 = time.time()
        self.ws.Range("A3").Value = 500000
        self.ws.Range("B3").Formula = "=BWPVNDUPPER(A3)"
        self.session.excel.CalculateFull()

        expected = "NĂM TRĂM NGHÌN ĐỒNG CHẴN."
        self.assertEqual(self.ws.Range("B3").Value2, expected)
        record_result("Tier 2", "Canonical =BWPVNDUPPER(A3) formula", True, duration=time.time() - t0)

    def test_convenience_aliases(self):
        """=VND(), =VNWORDS(), =VNDUPPER() produce identical results to canonical UDFs."""
        t0 = time.time()
        self.ws.Range("A1").Value = 125430000
        self.ws.Range("A2").Value = 125.05
        self.ws.Range("A3").Value = 500000

        self.ws.Range("B1").Formula = "=VND(A1)"
        self.ws.Range("B2").Formula = "=VNWORDS(A2)"
        self.ws.Range("B3").Formula = "=VNDUPPER(A3)"
        self.session.excel.CalculateFull()

        self.assertEqual(self.ws.Range("B1").Value2, "Một trăm hai mươi lăm triệu bốn trăm ba mươi nghìn đồng chẵn.")
        self.assertEqual(self.ws.Range("B2").Value2, "Một trăm hai mươi lăm phẩy không năm.")
        self.assertEqual(self.ws.Range("B3").Value2, "NĂM TRĂM NGHÌN ĐỒNG CHẴN.")
        record_result("Tier 2", "Convenience aliases =VND, =VNWORDS, =VNDUPPER", True, duration=time.time() - t0)

    def test_blank_cell_returns_empty_string(self):
        """Blank cell input produces empty string output without error."""
        t0 = time.time()
        self.ws.Range("A4").Value = None
        self.ws.Range("B4").Formula = "=BWPVND(A4)"
        self.ws.Range("C4").Formula = "=BWPVNWORDS(A4)"
        self.ws.Range("D4").Formula = "=BWPVNDUPPER(A4)"
        self.session.excel.CalculateFull()

        self.assertEqual(self.ws.Range("B4").Value2, "")
        self.assertEqual(self.ws.Range("C4").Value2, "")
        self.assertEqual(self.ws.Range("D4").Value2, "")
        record_result("Tier 2", "Blank cell returns empty string", True, duration=time.time() - t0)

    def test_overflow_and_invalid_input_return_xl_err_value(self):
        """Overflow (>999T), invalid text, and multi-cell argument return #VALUE!."""
        t0 = time.time()
        # Overflow
        self.ws.Range("A5").Value = 1000000000000000
        self.ws.Range("B5").Formula = "=BWPVND(A5)"
        self.ws.Range("C5").Formula = "=BWPVNWORDS(A5)"

        # Invalid text
        self.ws.Range("A6").Value = "invalid_string"
        self.ws.Range("B6").Formula = "=BWPVND(A6)"

        # Multi-cell range
        self.ws.Range("A7:A9").Value = [[10], [20], [30]]
        self.ws.Range("B7").Formula = "=BWPVND(A7:A9)"

        self.session.excel.CalculateFull()

        self.assertIn(self.ws.Range("B5").Value2, XL_ERR_VALUE_CODES)
        self.assertEqual(self.ws.Range("B5").Text, "#VALUE!")
        self.assertIn(self.ws.Range("C5").Value2, XL_ERR_VALUE_CODES)
        self.assertIn(self.ws.Range("B6").Value2, XL_ERR_VALUE_CODES)
        self.assertIn(self.ws.Range("B7").Value2, XL_ERR_VALUE_CODES)

        record_result("Tier 2", "Error conditions return #VALUE! (overflow, text, array)", True, duration=time.time() - t0)

    def test_udf_optional_parameters(self):
        """Verify optional parameters (AddChan=False, dialect styles)."""
        t0 = time.time()
        self.ws.Range("A1").Value = 125430000
        # AddChan = False: excludes "chẵn"
        self.ws.Range("B1").Formula = "=BWPVND(A1, FALSE)"
        # Dialect: ZeroStyle=1 (linh), ThousandStyle=1 (ngàn)
        self.ws.Range("A2").Value = 1005
        self.ws.Range("B2").Formula = "=BWPVND(A2, TRUE, 1, 1)"
        self.session.excel.CalculateFull()

        self.assertEqual(self.ws.Range("B1").Value2, "Một trăm hai mươi lăm triệu bốn trăm ba mươi nghìn đồng.")
        self.assertIn("linh", self.ws.Range("B2").Value2)
        self.assertIn("ngàn", self.ws.Range("B2").Value2)
        record_result("Tier 2", "UDF optional parameters (AddChan, ZeroStyle, ThousandStyle)", True, duration=time.time() - t0)


# ---------------------------------------------------------------------------
# Tier 3: Excel Integration Tests
# ---------------------------------------------------------------------------

class Tier3ExcelIntegrationTests(unittest.TestCase):
    """
    Tier 3: Excel Integration Tests:
    - Quick Convert via OnQuickConvertClick on 1D range
    - Destination auto-expansion
    - Cross-sheet conversion and same-sheet overlap rejection
    - 2-phase Undo execution and formula restoration
    - Formula-safe skip preservation
    """

    session: Optional[ExcelSession] = None
    addin_wb: Any = None
    addin_name: str = ""
    test_wb: Any = None
    ws: Any = None

    @classmethod
    def setUpClass(cls):
        target = Path(CURRENT_XLAM_PATH).resolve()
        if not target.is_file():
            raise FileNotFoundError(f"Target add-in not found: {target}")

        # Temporarily disable ShowBatchSummary in registry so OnQuickConvertClick executes headlessly without modal dialog
        cls._orig_summary = None
        try:
            reg_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\VB and VBA Program Settings\BWPConvertTTNVN\App")
            try:
                cls._orig_summary, _ = winreg.QueryValueEx(reg_key, "ShowBatchSummary")
            except FileNotFoundError:
                cls._orig_summary = None
            winreg.SetValueEx(reg_key, "ShowBatchSummary", 0, winreg.REG_SZ, "0")
            winreg.CloseKey(reg_key)
        except Exception as e:
            print(f"[WARN] Failed to adjust ShowBatchSummary registry key: {e}")

        cls.session = ExcelSession(visible=False, display_alerts=False)
        cls.session.__enter__()
        cls.addin_wb = cls.session.excel.Workbooks.Open(str(target))
        cls.addin_name = cls.addin_wb.Name
        cls.test_wb = cls.session.excel.Workbooks.Add()
        cls.ws = cls.test_wb.Worksheets(1)

    @classmethod
    def tearDownClass(cls):
        # Restore registry setting
        try:
            reg_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\VB and VBA Program Settings\BWPConvertTTNVN\App")
            if cls._orig_summary is not None:
                winreg.SetValueEx(reg_key, "ShowBatchSummary", 0, winreg.REG_SZ, str(cls._orig_summary))
            else:
                try:
                    winreg.DeleteValue(reg_key, "ShowBatchSummary")
                except Exception:
                    pass
            winreg.CloseKey(reg_key)
        except Exception:
            pass

        if cls.test_wb:
            try:
                cls.test_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if cls.addin_wb:
            try:
                cls.addin_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if cls.session:
            cls.session.__exit__(None, None, None)

    def setUp(self):
        try:
            self.ws.Activate()
        except Exception:
            pass
        self.ws.Cells.Clear()
        try:
            self.session.excel.Run(f"'{self.addin_name}'!UndoManager.ClearUndo")
        except Exception:
            pass

    def test_quick_convert_1d_range(self):
        """Quick Convert via OnQuickConvertClick on 1D vertical selection converts adjacent range."""
        t0 = time.time()
        self.ws.Activate()
        values = [[10000], [20000], [30000], [40000], [50000]]
        self.ws.Range("A1:A5").Value = values
        self.ws.Range("A1:A5").Select()

        # Pass ws COM object as control parameter
        macro = f"'{self.addin_name}'!RibbonCallbacks.OnQuickConvertClick"
        self.session.excel.Run(macro, self.ws)

        self.assertEqual(self.ws.Range("A1").Value2, "Mười nghìn đồng chẵn.")
        self.assertEqual(self.ws.Range("A3").Value2, "Ba mươi nghìn đồng chẵn.")
        self.assertEqual(self.ws.Range("A5").Value2, "Năm mươi nghìn đồng chẵn.")
        record_result("Tier 3", "Quick Convert via OnQuickConvertClick on 1D range (in-place)", True, duration=time.time() - t0)

    def test_destination_auto_expansion(self):
        """1D source range auto-expands single anchor cell to matching dimensions."""
        t0 = time.time()
        values = [[10000], [20000], [30000], [40000], [50000]]
        self.ws.Range("A1:A5").Value = values

        macro = f"'{self.addin_name}'!CellProcessor.ConvertRangeBridge"
        res = self.session.excel.Run(macro, self.ws.Range("A1:A5"), self.ws.Range("B1"))
        success, converted, skipped, errors, undo_avail, err_msg = res

        self.assertTrue(success, f"ConvertRangeBridge failed: {err_msg}")
        self.assertEqual(converted, 5)
        self.assertEqual(skipped, 0)
        self.assertEqual(errors, 0)
        self.assertTrue(undo_avail)
        self.assertEqual(self.ws.Range("B1").Value2, "Mười nghìn đồng chẵn.")
        self.assertEqual(self.ws.Range("B5").Value2, "Năm mươi nghìn đồng chẵn.")
        record_result("Tier 3", "Destination auto-expansion from single anchor cell", True, duration=time.time() - t0)

    def test_cross_sheet_and_same_sheet_overlap_guard(self):
        """Cross-sheet conversion succeeds; same-sheet overlap is rejected."""
        t0 = time.time()
        macro = f"'{self.addin_name}'!CellProcessor.ConvertRangeBridge"

        # 1. Cross-sheet conversion
        ws2 = self.test_wb.Worksheets.Add()
        ws2.Cells.Clear()
        self.ws.Range("A1:A3").Value = [[100], [200], [300]]
        res_cross = self.session.excel.Run(macro, self.ws.Range("A1:A3"), ws2.Range("B1"))
        self.assertTrue(res_cross[0], f"Cross-sheet failed: {res_cross[5]}")
        self.assertEqual(res_cross[1], 3)
        self.assertEqual(ws2.Range("B1").Value2, "Một trăm đồng chẵn.")

        # 2. Same-sheet overlap rejection
        res_overlap = self.session.excel.Run(macro, self.ws.Range("A1:A10"), self.ws.Range("A5:A14"))
        self.assertFalse(res_overlap[0], "Overlapping range must be rejected")
        self.assertEqual(res_overlap[1], 0)
        err_lower = str(res_overlap[5]).lower()
        self.assertTrue(
            "overlap" in err_lower,
            f"Expected overlap error, got: {res_overlap[5]}",
        )
        record_result("Tier 3", "Cross-sheet conversion and same-sheet overlap rejection", True, duration=time.time() - t0)

    def test_transactional_undo_and_formula_restoration(self):
        """2-phase Undo restores original formulas, values, and clears empty cells."""
        t0 = time.time()
        bridge = f"'{self.addin_name}'!CellProcessor.ConvertRangeBridge"
        undo_macro = f"'{self.addin_name}'!UndoManager.ExecuteUndo"

        # Setup destination: B1 formula, B2 constant, B3 empty
        self.ws.Range("A1").Value = 500
        self.ws.Range("A2").Value = 1000
        self.ws.Range("A3").Value = 2000
        self.ws.Range("B1").Formula = "=A1*2"
        self.ws.Range("B2").Value = "Old Constant"
        self.ws.Range("B3").Value = None

        # Convert
        res = self.session.excel.Run(bridge, self.ws.Range("A1:A3"), self.ws.Range("B1:B3"))
        self.assertTrue(res[0], f"Batch conversion failed: {res[5]}")
        self.assertEqual(self.ws.Range("B1").Value2, "Năm trăm đồng chẵn.")

        # Execute Undo
        undo_res = self.session.excel.Run(undo_macro)
        self.assertTrue(undo_res, "ExecuteUndo returned False")

        # Verify restoration
        self.assertTrue(self.ws.Range("B1").HasFormula)
        self.assertEqual(self.ws.Range("B1").Formula, "=A1*2")
        self.assertEqual(self.ws.Range("B2").Value2, "Old Constant")
        self.assertIsNone(self.ws.Range("B3").Value2)
        record_result("Tier 3", "2-Phase Undo execution and formula restoration", True, duration=time.time() - t0)

    def test_formula_safe_skip_preservation(self):
        """Blank source cells are skipped, leaving existing destination formulas untouched."""
        t0 = time.time()
        bridge = f"'{self.addin_name}'!CellProcessor.ConvertRangeBridge"

        self.ws.Range("A1").Value = 100
        self.ws.Range("A2").Value = None  # Blank source
        self.ws.Range("B1").Value = None
        self.ws.Range("B2").Formula = "=SUM(C1:C10)"

        res = self.session.excel.Run(bridge, self.ws.Range("A1:A2"), self.ws.Range("B1:B2"))
        success, converted, skipped, errors, undo_avail, err_msg = res

        self.assertTrue(success, f"ConvertRangeBridge failed: {err_msg}")
        self.assertEqual(converted, 1)
        self.assertEqual(skipped, 1)
        self.assertEqual(self.ws.Range("B1").Value2, "Một trăm đồng chẵn.")
        self.assertTrue(self.ws.Range("B2").HasFormula, "Destination formula was overwritten")
        self.assertEqual(self.ws.Range("B2").Formula, "=SUM(C1:C10)")
        record_result("Tier 3", "Formula-safe skip preservation for existing formulas", True, duration=time.time() - t0)


# ---------------------------------------------------------------------------
# Report Generator
# ---------------------------------------------------------------------------

def generate_test_report(
    xlam_path: Path,
    report_path: Path,
    records: List[TestRecord],
    total_elapsed: float,
) -> str:
    """
    Formats the complete test report and writes it to report_path.
    """
    total_tests = len(records)
    passed_tests = sum(1 for r in records if r.passed)
    failed_tests = total_tests - passed_tests

    # Calculate SHA-256 and size of target add-in
    sha256_hash = "N/A"
    file_size = 0
    if xlam_path.is_file():
        file_size = xlam_path.stat().st_size
        sha256_hash = hashlib.sha256(xlam_path.read_bytes()).hexdigest()

    lines = []
    lines.append("=" * 72)
    lines.append("  BWPConvertTTNVN - Automated 3-Tier Test Report")
    lines.append("=" * 72)
    lines.append(f"Timestamp:       {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"Target Add-in:   {xlam_path}")
    lines.append(f"File Size:       {file_size:,} bytes")
    lines.append(f"SHA-256 Hash:    {sha256_hash}")
    lines.append(f"Duration:        {total_elapsed:.2f}s")
    lines.append(f"Overall Result:  {'ALL TESTS PASSED' if failed_tests == 0 else 'TESTS FAILED'}")
    lines.append(f"Summary:         {total_tests} total | {passed_tests} passed | {failed_tests} failed")
    lines.append("-" * 72)

    for tier_name in ["Tier 1", "Tier 2", "Tier 3"]:
        tier_records = [r for r in records if r.tier == tier_name]
        lines.append(f"\n[{tier_name} Tests] ({len(tier_records)} items)")
        lines.append("-" * 72)
        for r in tier_records:
            status = "[ PASS ]" if r.passed else "[ FAIL ]"
            dur = f"({r.duration:.3f}s)" if r.duration > 0 else ""
            lines.append(f"  {status} {r.name} {dur}")
            if r.detail:
                lines.append(f"           Detail: {r.detail}")

    lines.append("\n" + "=" * 72)
    if failed_tests == 0:
        lines.append("  [SUCCESS] All 3 tiers passed successfully against packaged .xlam!")
    else:
        lines.append(f"  [FAILURE] {failed_tests} test(s) failed.")
    lines.append("=" * 72 + "\n")

    report_text = "\n".join(lines)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(report_text, encoding="utf-8")
    return report_text


# ---------------------------------------------------------------------------
# Runner Entry Point
# ---------------------------------------------------------------------------

def run_all_tests(
    xlam_path: Optional[Union[str, Path]] = None,
    report_path: Optional[Union[str, Path]] = None,
    tier_filter: Optional[str] = None,
) -> Tuple[bool, str]:
    """
    Executes all 3 test tiers and produces a formatted test report.
    """
    global CURRENT_XLAM_PATH, GLOBAL_TEST_RECORDS
    if xlam_path:
        CURRENT_XLAM_PATH = Path(xlam_path).resolve()
    target_report = Path(report_path).resolve() if report_path else DEFAULT_REPORT_PATH
    GLOBAL_TEST_RECORDS.clear()

    if not CURRENT_XLAM_PATH.is_file():
        raise FileNotFoundError(f"Target add-in does not exist: {CURRENT_XLAM_PATH}")

    suite = unittest.TestSuite()
    loader = unittest.TestLoader()

    if tier_filter in (None, "all", "1"):
        suite.addTests(loader.loadTestsFromTestCase(Tier1CoreApiTests))
    if tier_filter in (None, "all", "2"):
        suite.addTests(loader.loadTestsFromTestCase(Tier2WorksheetUdfTests))
    if tier_filter in (None, "all", "3"):
        suite.addTests(loader.loadTestsFromTestCase(Tier3ExcelIntegrationTests))

    start_time = time.time()
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    elapsed = time.time() - start_time

    # Record any failures or errors that unittest caught
    for test, err in result.failures:
        record_result("Unknown", str(test), False, f"Failure: {err}")
    for test, err in result.errors:
        record_result("Unknown", str(test), False, f"Error: {err}")

    success = result.wasSuccessful()
    report_text = generate_test_report(CURRENT_XLAM_PATH, target_report, GLOBAL_TEST_RECORDS, elapsed)
    print("\n" + report_text)
    return success, report_text


def main():
    parser = argparse.ArgumentParser(description="BWPConvertTTNVN 3-Tier Automated Test Runner")
    parser.add_argument("--xlam", default=str(DEFAULT_XLAM_PATH), help=f"Path to .xlam file (default: {DEFAULT_XLAM_PATH})")
    parser.add_argument("--report", default=str(DEFAULT_REPORT_PATH), help=f"Path to output report file (default: {DEFAULT_REPORT_PATH})")
    parser.add_argument("--tier", choices=["all", "1", "2", "3"], default="all", help="Filter by tier (default: all)")

    args = parser.parse_args()

    try:
        success, _ = run_all_tests(
            xlam_path=args.xlam,
            report_path=args.report,
            tier_filter=args.tier,
        )
        sys.exit(0 if success else 1)
    except Exception as e:
        print(f"\n[TEST RUNNER ERROR] {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
