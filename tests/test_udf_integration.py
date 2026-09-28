"""
Automated integration test suite for UDF.bas.
Validates:
1. Pure 7-bit ASCII safety of src/excel/UDF.bas.
2. Live Excel COM automation testing of worksheet formulas:
   - Canonical =BWPVND(A1) produces sentence-case Vietnamese Dong with "dong chan."
   - Canonical =BWPVNWORDS(A2) reads general numbers with decimal digits ("phay khong nam.")
   - Canonical =BWPVNDUPPER(A3) produces all-uppercase Vietnamese Dong.
   - Blank/empty input cell returns "" (empty string).
   - Invalid non-numeric input returns #VALUE! (-2146826281).
   - Overflow input (> 999T) returns #VALUE!.
   - Multi-cell range or array argument returns #VALUE!.
   - Aliases =VND(), =VNWORDS(), =VNDUPPER() produce identical outputs.
   - Optional parameters (AddChan, ZeroStyle, ThousandStyle).
   - Registry isolation: changing Registry settings has zero effect on formula outputs.
"""

import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from scripts.build import ExcelSession

XL_ERR_VALUE = -2146826273  # Excel xlErrValue (2015 / 0x800A07DF)
XL_ERR_VALUE_CODES = (-2146826273, -2146826281)


class TestUdfAsciiSafety(unittest.TestCase):
    """Verifies that UDF.bas in src/excel/ is strictly 7-bit ASCII safe."""

    def test_udf_bas_is_pure_ascii(self):
        udf_file = PROJECT_ROOT / "src" / "excel" / "UDF.bas"
        self.assertTrue(udf_file.is_file(), "src/excel/UDF.bas must exist")

        content = udf_file.read_bytes()
        non_ascii = [b for b in content if b > 127]
        self.assertEqual(
            len(non_ascii),
            0,
            f"Found {len(non_ascii)} non-ASCII bytes in UDF.bas: {non_ascii[:10]}",
        )


class TestUdfIntegrationLiveCOM(unittest.TestCase):
    """Executes worksheet UDF tests via live Excel COM automation."""

    excel_session = None
    wb = None
    wb_name = None
    ws = None

    @classmethod
    def setUpClass(cls):
        cls.excel_session = ExcelSession(visible=False, display_alerts=False)
        cls.excel_session.__enter__()
        cls.excel = cls.excel_session.excel
        cls.wb = cls.excel.Workbooks.Add()
        cls.wb_name = cls.wb.Name
        cls.ws = cls.wb.Worksheets(1)

        # Import modules in dependency order
        modules_to_import = [
            ("src/core", "CoreTypes.bas"),
            ("src/core", "UnicodeText.bas"),
            ("src/core", "VietnameseNumber.bas"),
            ("src/core", "Currency.bas"),
            ("src/core", "TextFormatter.bas"),
            ("src/core", "CoreCoordinator.bas"),
            ("src/excel", "Settings.bas"),
            ("src/excel", "UDF.bas"),
        ]

        for sub_dir, mod_name in modules_to_import:
            mod_path = PROJECT_ROOT / sub_dir / mod_name
            cls.wb.VBProject.VBComponents.Import(str(mod_path))

    @classmethod
    def tearDownClass(cls):
        # Reset any registry settings modified during testing
        try:
            reset_macro = f"'{cls.wb_name}'!Settings.ResetSettingsBridge"
            cls.excel.Run(reset_macro)
        except Exception:
            pass

        if cls.wb:
            try:
                cls.wb.Close(SaveChanges=False)
            except Exception:
                pass
        if cls.excel_session:
            cls.excel_session.__exit__(None, None, None)

    def setUp(self):
        self.ws.Cells.Clear()

    # --------------------------------------------------------------------------
    # 1. Canonical Worksheet Formulas (Brief Requirements 44-46)
    # --------------------------------------------------------------------------

    def test_canonical_bwpvnd_formula(self):
        """A1 = 125430000, B1 = '=BWPVND(A1)' -> 'Một trăm hai mươi lăm triệu bốn trăm ba mươi nghìn đồng chẵn.'"""
        self.ws.Range("A1").Value = 125430000
        self.ws.Range("B1").Formula = "=BWPVND(A1)"
        self.excel.CalculateFull()

        expected = "Một trăm hai mươi lăm triệu bốn trăm ba mươi nghìn đồng chẵn."
        self.assertEqual(self.ws.Range("B1").Value2, expected)

    def test_canonical_bwpvnwords_decimal_formula(self):
        """A2 = 125.05, B2 = '=BWPVNWORDS(A2)' -> 'Một trăm hai mươi lăm phẩy không năm.'"""
        self.ws.Range("A2").Value = 125.05
        self.ws.Range("B2").Formula = "=BWPVNWORDS(A2)"
        self.excel.CalculateFull()

        expected = "Một trăm hai mươi lăm phẩy không năm."
        self.assertEqual(self.ws.Range("B2").Value2, expected)

    def test_canonical_bwpvndupper_formula(self):
        """A3 = 500000, B3 = '=BWPVNDUPPER(A3)' -> 'NĂM TRĂM NGHÌN ĐỒNG CHẴN.'"""
        self.ws.Range("A3").Value = 500000
        self.ws.Range("B3").Formula = "=BWPVNDUPPER(A3)"
        self.excel.CalculateFull()

        expected = "NĂM TRĂM NGHÌN ĐỒNG CHẴN."
        self.assertEqual(self.ws.Range("B3").Value2, expected)

    # --------------------------------------------------------------------------
    # 2. Blank and Empty Cell Handling (Brief Requirement 47)
    # --------------------------------------------------------------------------

    def test_blank_cell_returns_empty_string(self):
        """A4 = '' (blank cell), B4 = '=BWPVND(A4)' -> B4.Value2 == ''"""
        self.ws.Range("A4").Value = None  # Truly blank cell
        self.ws.Range("B4").Formula = "=BWPVND(A4)"
        self.ws.Range("C4").Formula = "=BWPVNWORDS(A4)"
        self.ws.Range("D4").Formula = "=BWPVNDUPPER(A4)"
        self.excel.CalculateFull()

        self.assertEqual(self.ws.Range("B4").Value2, "")
        self.assertEqual(self.ws.Range("C4").Value2, "")
        self.assertEqual(self.ws.Range("D4").Value2, "")

    def test_empty_string_cell_returns_empty_string(self):
        """A4 = '\"\"' (cell with empty string), formula returns empty string."""
        self.ws.Range("A4").Formula = '=""'
        self.ws.Range("B4").Formula = "=BWPVND(A4)"
        self.excel.CalculateFull()

        self.assertEqual(self.ws.Range("B4").Value2, "")

    # --------------------------------------------------------------------------
    # 3. Invalid Input and Error Handling (Brief Requirements 48-50)
    # --------------------------------------------------------------------------

    def test_invalid_text_returns_xl_err_value(self):
        """A5 = 'invalid', B5 = '=BWPVND(A5)' -> error #VALUE! (-2146826281)."""
        self.ws.Range("A5").Value = "invalid"
        self.ws.Range("B5").Formula = "=BWPVND(A5)"
        self.excel.CalculateFull()

        self.assertEqual(self.ws.Range("B5").Value2, XL_ERR_VALUE)
        self.assertEqual(self.ws.Range("B5").Text, "#VALUE!")

    def test_overflow_returns_xl_err_value(self):
        """A6 = 1000000000000000 (overflow > 999T), B6 = '=BWPVND(A6)' -> #VALUE!."""
        self.ws.Range("A6").Value = 1000000000000000
        self.ws.Range("B6").Formula = "=BWPVND(A6)"
        self.excel.CalculateFull()

        self.assertEqual(self.ws.Range("B6").Value2, XL_ERR_VALUE)
        self.assertEqual(self.ws.Range("B6").Text, "#VALUE!")

    def test_multicell_array_argument_returns_xl_err_value(self):
        """Multi-cell array argument =BWPVND(A1:A3) -> asserts #VALUE!."""
        self.ws.Range("A1:A3").Value = [[100], [200], [300]]
        self.ws.Range("B1").Formula = "=BWPVND(A1:A3)"
        self.ws.Range("B2").Formula = "=BWPVNWORDS(A1:A3)"
        self.ws.Range("B3").Formula = "=BWPVNDUPPER(A1:A3)"
        self.excel.CalculateFull()

        self.assertEqual(self.ws.Range("B1").Value2, XL_ERR_VALUE)
        self.assertEqual(self.ws.Range("B1").Text, "#VALUE!")
        self.assertEqual(self.ws.Range("B2").Value2, XL_ERR_VALUE)
        self.assertEqual(self.ws.Range("B3").Value2, XL_ERR_VALUE)

    def test_array_constant_argument_returns_xl_err_value(self):
        """Array constant argument =BWPVND({1, 2, 3}) -> asserts #VALUE!."""
        self.ws.Range("B1").Formula = "=BWPVND({1, 2, 3})"
        self.excel.CalculateFull()

        self.assertEqual(self.ws.Range("B1").Value2, XL_ERR_VALUE)

    def test_precedent_error_cell_returns_xl_err_value(self):
        """Precedent cell with error (e.g. =1/0) causes UDF to return #VALUE!."""
        self.ws.Range("A1").Formula = "=1/0"
        self.ws.Range("B1").Formula = "=BWPVND(A1)"
        self.excel.CalculateFull()

        self.assertEqual(self.ws.Range("B1").Value2, XL_ERR_VALUE)

    # --------------------------------------------------------------------------
    # 4. Convenience Aliases (Brief Requirement 51)
    # --------------------------------------------------------------------------

    def test_convenience_aliases_match_canonical(self):
        """Aliases =VND(A1), =VNWORDS(A2), =VNDUPPER(A3) work identically to canonicals."""
        self.ws.Range("A1").Value = 125430000
        self.ws.Range("A2").Value = 125.05
        self.ws.Range("A3").Value = 500000

        self.ws.Range("B1").Formula = "=VND(A1)"
        self.ws.Range("B2").Formula = "=VNWORDS(A2)"
        self.ws.Range("B3").Formula = "=VNDUPPER(A3)"
        self.excel.CalculateFull()

        self.assertEqual(
            self.ws.Range("B1").Value2,
            "Một trăm hai mươi lăm triệu bốn trăm ba mươi nghìn đồng chẵn.",
        )
        self.assertEqual(
            self.ws.Range("B2").Value2,
            "Một trăm hai mươi lăm phẩy không năm.",
        )
        self.assertEqual(
            self.ws.Range("B3").Value2,
            "NĂM TRĂM NGHÌN ĐỒNG CHẴN.",
        )

    # --------------------------------------------------------------------------
    # 5. Optional Parameter Variations
    # --------------------------------------------------------------------------

    def test_add_chan_false_variations(self):
        """Verify AddChan=False omits 'chẵn' in both sentence case and uppercase."""
        self.ws.Range("A1").Value = 125430000
        self.ws.Range("A2").Value = 500000

        self.ws.Range("B1").Formula = "=BWPVND(A1, FALSE)"
        self.ws.Range("B2").Formula = "=BWPVNDUPPER(A2, FALSE)"
        self.excel.CalculateFull()

        self.assertEqual(
            self.ws.Range("B1").Value2,
            "Một trăm hai mươi lăm triệu bốn trăm ba mươi nghìn đồng.",
        )
        self.assertEqual(
            self.ws.Range("B2").Value2,
            "NĂM TRĂM NGHÌN ĐỒNG.",
        )

    def test_zero_and_thousand_style_parameters(self):
        """Verify ZeroStyle (0=le, 1=linh) and ThousandStyle (0=nghin, 1=ngan)."""
        self.ws.Range("A1").Value = 1005000

        # ZeroStyle=1 (linh), ThousandStyle=1 (ngan)
        self.ws.Range("B1").Formula = "=BWPVND(A1, TRUE, 1, 1)"
        self.ws.Range("B2").Formula = "=BWPVNWORDS(A1, 1, 1)"
        self.excel.CalculateFull()

        self.assertEqual(
            self.ws.Range("B1").Value2,
            "Một triệu không trăm linh năm ngàn đồng chẵn.",
        )
        self.assertEqual(
            self.ws.Range("B2").Value2,
            "Một triệu không trăm linh năm ngàn.",
        )

    def test_invalid_optional_parameter_returns_xl_err_value(self):
        """Invalid ZeroStyle (e.g. 5) or ThousandStyle returns #VALUE!."""
        self.ws.Range("A1").Value = 1000
        self.ws.Range("B1").Formula = "=BWPVND(A1, TRUE, 5, 0)"
        self.ws.Range("B2").Formula = "=BWPVNWORDS(A1, 0, 9)"
        self.ws.Range("B3").Formula = '=BWPVND(A1, "not_a_bool")'
        self.excel.CalculateFull()

        self.assertEqual(self.ws.Range("B1").Value2, XL_ERR_VALUE)
        self.assertEqual(self.ws.Range("B2").Value2, XL_ERR_VALUE)
        self.assertEqual(self.ws.Range("B3").Value2, XL_ERR_VALUE)

    def test_bwpvnd_truncates_decimal_toward_zero(self):
        """=BWPVND(125.75) truncates toward zero (Fix) -> 'Một trăm hai mươi lăm đồng chẵn.'"""
        self.ws.Range("A1").Value = 125.75
        self.ws.Range("B1").Formula = "=BWPVND(A1)"
        self.excel.CalculateFull()

        self.assertEqual(self.ws.Range("B1").Value2, "Một trăm hai mươi lăm đồng chẵn.")

    # --------------------------------------------------------------------------
    # 6. Registry Isolation Verification (Brief Requirement 52)
    # --------------------------------------------------------------------------

    def test_registry_isolation_formula_outputs_remain_deterministic(self):
        """
        Verify that modifying Registry settings does NOT alter formula outputs.
        UDFs must be 100% deterministic from formula arguments alone.
        """
        save_macro = f"'{self.wb_name}'!Settings.SaveSettingsBridge"
        reset_macro = f"'{self.wb_name}'!Settings.ResetSettingsBridge"

        # 1. Setup standard formula cells
        self.ws.Range("A1").Value = 125430000
        self.ws.Range("A2").Value = 125.05
        self.ws.Range("A3").Value = 500000

        self.ws.Range("B1").Formula = "=BWPVND(A1)"
        self.ws.Range("B2").Formula = "=BWPVNWORDS(A2)"
        self.ws.Range("B3").Formula = "=BWPVNDUPPER(A3)"
        self.excel.CalculateFull()

        expected_b1 = "Một trăm hai mươi lăm triệu bốn trăm ba mươi nghìn đồng chẵn."
        expected_b2 = "Một trăm hai mươi lăm phẩy không năm."
        expected_b3 = "NĂM TRĂM NGHÌN ĐỒNG CHẴN."

        self.assertEqual(self.ws.Range("B1").Value2, expected_b1)
        self.assertEqual(self.ws.Range("B2").Value2, expected_b2)
        self.assertEqual(self.ws.Range("B3").Value2, expected_b3)

        # 2. Modify Registry to polar opposites:
        # Zero=1 (linh), Thousand=1 (ngan), DecMode=0 (ignore), AddChan=False,
        # Casing=2 (lower), AddPeriod=False
        self.excel.Run(
            save_macro,
            1, 1, 1, 0,
            0, False, 0, "", "", "",
            2, False,
            0, 0, True, True, 1,
        )

        # 3. Force full recalculation
        self.excel.CalculateFull()

        # 4. Assert outputs are completely UNCHANGED
        self.assertEqual(
            self.ws.Range("B1").Value2,
            expected_b1,
            "UDF BWPVND output changed after registry modification!",
        )
        self.assertEqual(
            self.ws.Range("B2").Value2,
            expected_b2,
            "UDF BWPVNWORDS output changed after registry modification!",
        )
        self.assertEqual(
            self.ws.Range("B3").Value2,
            expected_b3,
            "UDF BWPVNDUPPER output changed after registry modification!",
        )

        # 5. Clean up registry
        self.excel.Run(reset_macro)


if __name__ == "__main__":
    unittest.main()
