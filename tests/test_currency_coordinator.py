"""
Automated test suite for Currency.bas, TextFormatter.bas, and CoreCoordinator.bas.
Validates:
1. Pure 7-bit ASCII safety of all .bas files in src/core/.
2. Deterministic half-away-from-zero rounding in pure VBA.
3. Currency formatting, sub-units, carry rounding, and sign handling.
4. Whitespace collapsing, Unicode casing, and trailing punctuation normalization.
5. Live Excel COM automation testing of the complete CoreCoordinator pipeline.
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


class TestCoreAsciiSafety(unittest.TestCase):
    """Verifies that all VBA source modules in src/core/ are strictly 7-bit ASCII safe."""

    def test_all_core_bas_files_are_ascii(self):
        core_dir = PROJECT_ROOT / "src" / "core"
        bas_files = sorted(core_dir.glob("*.bas"))
        self.assertGreaterEqual(len(bas_files), 6, "Expected at least 6 core .bas files")

        for bas_file in bas_files:
            with self.subTest(file=bas_file.name):
                content = bas_file.read_bytes()
                non_ascii = [b for b in content if b > 127]
                self.assertEqual(
                    len(non_ascii),
                    0,
                    f"Found {len(non_ascii)} non-ASCII bytes in {bas_file.name}: {non_ascii[:10]}",
                )


class TestCurrencyCoordinatorLiveCOM(unittest.TestCase):
    """Executes currency and coordinator tests via live Excel COM automation."""

    excel_session = None
    wb = None
    wb_name = None

    @classmethod
    def setUpClass(cls):
        cls.excel_session = ExcelSession(visible=False, display_alerts=False)
        cls.excel_session.__enter__()
        cls.excel = cls.excel_session.excel
        cls.wb = cls.excel.Workbooks.Add()
        cls.wb_name = cls.wb.Name

        core_modules = [
            "CoreTypes.bas",
            "UnicodeText.bas",
            "VietnameseNumber.bas",
            "Currency.bas",
            "TextFormatter.bas",
            "CoreCoordinator.bas",
        ]

        for mod_name in core_modules:
            mod_path = PROJECT_ROOT / "src" / "core" / mod_name
            cls.wb.VBProject.VBComponents.Import(str(mod_path))

    @classmethod
    def tearDownClass(cls):
        if cls.wb:
            cls.wb.Close(SaveChanges=False)
        if cls.excel_session:
            cls.excel_session.__exit__(None, None, None)

    # --------------------------------------------------------------------------
    # 1. Deterministic Half-Away-From-Zero Rounding
    # --------------------------------------------------------------------------

    def test_round_half_away_from_zero(self):
        """Test pure VBA RoundHalfAwayFromZero deterministic behavior."""
        round_macro = f"'{self.wb_name}'!VnCurrency.RoundHalfAwayFromZero"

        cases = [
            (1.005, 2, 1.01),
            (1.004, 2, 1.00),
            (-1.005, 2, -1.01),
            (-1.004, 2, -1.00),
            (1.999, 2, 2.00),
            (1.999, 0, 2.00),
            (-0.50, 2, -0.50),
            (0.0, 2, 0.0),
            (2.675, 2, 2.68),
            (0.025, 2, 0.03),
            (0.035, 2, 0.04),
            (0.045, 2, 0.05),
            (0.055, 2, 0.06),
            (0.065, 2, 0.07),
            (0.075, 2, 0.08),
            (0.085, 2, 0.09),
            (0.095, 2, 0.10),
        ]

        for val, dec_places, expected in cases:
            with self.subTest(val=val, dec_places=dec_places):
                actual = self.excel.Run(round_macro, val, dec_places)
                self.assertAlmostEqual(
                    float(actual),
                    expected,
                    places=4,
                    msg=f"RoundHalfAwayFromZero({val}, {dec_places}) failed: got {actual}, expected {expected}",
                )

    # --------------------------------------------------------------------------
    # 2. TextFormatter: Whitespace, Casing, and Punctuation
    # --------------------------------------------------------------------------

    def test_text_formatter_whitespace_and_punctuation(self):
        """Test TextFormatter whitespace collapsing and punctuation handling."""
        fmt_macro = f"'{self.wb_name}'!TextFormatter.FormatText"

        # Whitespace collapse & trim
        self.assertEqual(
            self.excel.Run(fmt_macro, "  một   trăm   đồng  ", 0, True),
            "Một trăm đồng.",
        )

        # Space before period
        self.assertEqual(
            self.excel.Run(fmt_macro, "một trăm đồng .", 0, True),
            "Một trăm đồng.",
        )

        # Ensure no double periods
        self.assertEqual(
            self.excel.Run(fmt_macro, "Một trăm đồng.", 0, True),
            "Một trăm đồng.",
        )

        # Empty string returns empty
        self.assertEqual(self.excel.Run(fmt_macro, "", 0, True), "")
        self.assertEqual(self.excel.Run(fmt_macro, "   ", 0, True), "")

    def test_text_formatter_casing_modes(self):
        """Test TextFormatter casing styles: Sentence, Upper, Lower."""
        fmt_macro = f"'{self.wb_name}'!TextFormatter.FormatText"
        sample = "một trăm hai mươi lăm nghìn đồng chẵn"

        # VnCaseSentence (0)
        self.assertEqual(
            self.excel.Run(fmt_macro, sample, 0, True),
            "Một trăm hai mươi lăm nghìn đồng chẵn.",
        )

        # VnCaseUpper (1)
        self.assertEqual(
            self.excel.Run(fmt_macro, sample, 1, True),
            "MỘT TRĂM HAI MƯƠI LĂM NGHÌN ĐỒNG CHẴN.",
        )

        # VnCaseLower (2)
        self.assertEqual(
            self.excel.Run(fmt_macro, sample, 2, True),
            "một trăm hai mươi lăm nghìn đồng chẵn.",
        )

    def test_text_formatter_trailing_period(self):
        """Test TextFormatter trailing period toggle (AddPeriod = True / False)."""
        fmt_macro = f"'{self.wb_name}'!TextFormatter.FormatText"
        sample = "một trăm nghìn đồng chẵn"

        # AddPeriod = True
        self.assertEqual(
            self.excel.Run(fmt_macro, sample, 0, True),
            "Một trăm nghìn đồng chẵn.",
        )

        # AddPeriod = False
        self.assertEqual(
            self.excel.Run(fmt_macro, sample, 0, False),
            "Một trăm nghìn đồng chẵn",
        )

    # --------------------------------------------------------------------------
    # 3. CoreCoordinator: Default VND Conversions
    # --------------------------------------------------------------------------

    def test_convert_number_default_standard_vnd(self):
        """Verify ConvertNumberDefault(125430000) produces exact expected VND words."""
        default_macro = f"'{self.wb_name}'!CoreCoordinator.ConvertNumberDefault"
        actual = self.excel.Run(default_macro, 125430000)
        expected = "Một trăm hai mươi lăm triệu bốn trăm ba mươi nghìn đồng chẵn."
        self.assertEqual(actual, expected)

    def test_convert_number_default_zero(self):
        """Verify ConvertNumberDefault(0) produces exact 'Không đồng chẵn.'."""
        default_macro = f"'{self.wb_name}'!CoreCoordinator.ConvertNumberDefault"
        actual = self.excel.Run(default_macro, 0)
        expected = "Không đồng chẵn."
        self.assertEqual(actual, expected)

    def test_convert_number_default_negative_vnd(self):
        """Verify negative VND -150000 produces exact 'Âm một trăm năm mươi nghìn đồng chẵn.'."""
        default_macro = f"'{self.wb_name}'!CoreCoordinator.ConvertNumberDefault"
        actual = self.excel.Run(default_macro, -150000)
        expected = "Âm một trăm năm mươi nghìn đồng chẵn."
        self.assertEqual(actual, expected)

    # --------------------------------------------------------------------------
    # 4. CoreCoordinator: USD Sub-units and Carry Rounding
    # --------------------------------------------------------------------------

    def test_usd_sub_units_one_cent(self):
        """Verify 1.005 USD rounds to 1.01 and produces 'Một đô la Mỹ một cent.'."""
        opt_macro = f"'{self.wb_name}'!CoreCoordinator.ConvertNumberWithOptions"
        # CurrType=1 (USD), AddChan=True, DecPlaces=2, Zero=0, Thousand=0, Four=0, DecMode=0, Casing=0, AddPeriod=True
        actual = self.excel.Run(opt_macro, 1.005, 1, True, 2, 0, 0, 0, 0, 0, True)
        expected = "Một đô la Mỹ một cent."
        self.assertEqual(actual, expected)

    def test_usd_sub_units_carry_rounding(self):
        """Verify 1.999 USD carry-rounds to 2.00 and produces 'Hai đô la Mỹ chẵn.'."""
        opt_macro = f"'{self.wb_name}'!CoreCoordinator.ConvertNumberWithOptions"
        actual = self.excel.Run(opt_macro, 1.999, 1, True, 2, 0, 0, 0, 0, 0, True)
        expected = "Hai đô la Mỹ chẵn."
        self.assertEqual(actual, expected)

    def test_usd_negative_sub_units(self):
        """Verify -0.50 USD produces 'Âm không đô la Mỹ năm mươi cent.'."""
        opt_macro = f"'{self.wb_name}'!CoreCoordinator.ConvertNumberWithOptions"
        actual = self.excel.Run(opt_macro, -0.50, 1, True, 2, 0, 0, 0, 0, 0, True)
        expected = "Âm không đô la Mỹ năm mươi cent."
        self.assertEqual(actual, expected)

    def test_usd_exact_amount_without_chan(self):
        """Verify USD integer amount with AddChan=False omits 'chẵn'."""
        opt_macro = f"'{self.wb_name}'!CoreCoordinator.ConvertNumberWithOptions"
        actual = self.excel.Run(opt_macro, 100, 1, False, 2, 0, 0, 0, 0, 0, True)
        expected = "Một trăm đô la Mỹ."
        self.assertEqual(actual, expected)

    def test_usd_various_cent_amounts(self):
        """Verify various USD cents readings (single digit, teens, tens)."""
        opt_macro = f"'{self.wb_name}'!CoreCoordinator.ConvertNumberWithOptions"

        cases = [
            (0.05, "Không đô la Mỹ năm cent."),
            (0.10, "Không đô la Mỹ mười cent."),
            (0.15, "Không đô la Mỹ mười lăm cent."),
            (0.21, "Không đô la Mỹ hai mươi mốt cent."),
            (0.24, "Không đô la Mỹ hai mươi tư cent."),
            (0.99, "Không đô la Mỹ chín mươi chín cent."),
        ]

        for val, expected in cases:
            with self.subTest(val=val):
                actual = self.excel.Run(opt_macro, val, 1, True, 2, 0, 0, 0, 0, 0, True)
                self.assertEqual(actual, expected)

    # --------------------------------------------------------------------------
    # 5. CoreCoordinator: Casing Modes and Period Control
    # --------------------------------------------------------------------------

    def test_coordinator_casing_modes(self):
        """Test ConvertNumberWithOptions casing modes: Sentence, Upper, Lower."""
        opt_macro = f"'{self.wb_name}'!CoreCoordinator.ConvertNumberWithOptions"

        # Sentence case (0)
        self.assertEqual(
            self.excel.Run(opt_macro, 125430000, 0, True, 0, 0, 0, 0, 0, 0, True),
            "Một trăm hai mươi lăm triệu bốn trăm ba mươi nghìn đồng chẵn.",
        )

        # Upper case (1)
        self.assertEqual(
            self.excel.Run(opt_macro, 125430000, 0, True, 0, 0, 0, 0, 0, 1, True),
            "MỘT TRĂM HAI MƯƠI LĂM TRIỆU BỐN TRĂM BA MƯƠI NGHÌN ĐỒNG CHẴN.",
        )

        # Lower case (2)
        self.assertEqual(
            self.excel.Run(opt_macro, 125430000, 0, True, 0, 0, 0, 0, 0, 2, True),
            "một trăm hai mươi lăm triệu bốn trăm ba mươi nghìn đồng chẵn.",
        )

    def test_coordinator_trailing_period_toggle(self):
        """Test ConvertNumberWithOptions AddPeriod toggle."""
        opt_macro = f"'{self.wb_name}'!CoreCoordinator.ConvertNumberWithOptions"

        with_dot = self.excel.Run(opt_macro, 50000, 0, True, 0, 0, 0, 0, 0, 0, True)
        self.assertEqual(with_dot, "Năm mươi nghìn đồng chẵn.")

        without_dot = self.excel.Run(opt_macro, 50000, 0, True, 0, 0, 0, 0, 0, 0, False)
        self.assertEqual(without_dot, "Năm mươi nghìn đồng chẵn")

    # --------------------------------------------------------------------------
    # 6. Currency Type: None and Custom
    # --------------------------------------------------------------------------

    def test_currency_none(self):
        """Test VnCurrNone (CurrType=2) reads base number without currency unit."""
        opt_macro = f"'{self.wb_name}'!CoreCoordinator.ConvertNumberWithOptions"
        actual = self.excel.Run(opt_macro, 125000, 2, False, 0, 0, 0, 0, 0, 0, True)
        self.assertEqual(actual, "Một trăm hai mươi lăm nghìn.")

    def test_currency_custom_via_ex(self):
        """Test custom currency formatting via NumberToCurrencyWordsEx."""
        ex_macro = f"'{self.wb_name}'!VnCurrency.NumberToCurrencyWordsEx"
        fmt_macro = f"'{self.wb_name}'!TextFormatter.FormatText"

        # Custom currency: 15.50 EUR with cents
        # Value=15.50, CurrType=3, AddChan=True, DecPlaces=2, Prefix="", Suffix="euro", SubUnit="cent"
        raw = self.excel.Run(
            ex_macro, 15.50, 3, True, 2, "", "euro", "cent", 0, 0, 0, 0
        )
        formatted = self.excel.Run(fmt_macro, raw, 0, True)
        self.assertEqual(formatted, "Mười lăm euro năm mươi cent.")

        # Custom currency: 100 EUR integer with AddChan=True
        raw_int = self.excel.Run(
            ex_macro, 100, 3, True, 2, "", "euro", "cent", 0, 0, 0, 0
        )
        formatted_int = self.excel.Run(fmt_macro, raw_int, 0, True)
        self.assertEqual(formatted_int, "Một trăm euro chẵn.")

    # --------------------------------------------------------------------------
    # 7. Error Handling and Boundary Guards
    # --------------------------------------------------------------------------

    def test_error_handling_invalid_input(self):
        """Verify invalid input is handled safely via TryConvertNumberWithOptions."""
        try_macro = f"'{self.wb_name}'!CoreCoordinator.TryConvertNumberWithOptions"

        # Non-numeric input
        res = self.excel.Run(try_macro, "abc", 0, True, 0, 0, 0, 0, 0, 0, True, "")
        self.assertEqual(res, "", "Expected empty string on invalid input")

        # Out of range input (exceeds MAX_SUPPORTED_VALUE)
        res_range = self.excel.Run(
            try_macro, 1000000000000000, 0, True, 0, 0, 0, 0, 0, 0, True, ""
        )
        self.assertEqual(res_range, "", "Expected empty string on out-of-range input")


if __name__ == "__main__":
    unittest.main()
