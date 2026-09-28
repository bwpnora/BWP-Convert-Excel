"""
Automated test suite for VietnameseNumber.bas and expected_cases.csv.
Validates pure 7-bit ASCII safety of core VBA code, in-memory VBA test assertions via RunCoreTests,
and live Excel COM execution of all 50+ test cases in expected_cases.csv driving the real VBA engine.
"""

import csv
import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from scripts.build import ExcelSession


class TestNumberEngineStatic(unittest.TestCase):
    """Static verification of ASCII safety and expected_cases.csv integrity."""

    def test_ascii_safety_vietnamese_number(self):
        """Verify VietnameseNumber.bas contains strictly 7-bit ASCII characters."""
        path = PROJECT_ROOT / "src" / "core" / "VietnameseNumber.bas"
        self.assertTrue(path.is_file(), f"VietnameseNumber.bas not found at {path}")
        content = path.read_bytes()
        non_ascii = [b for b in content if b > 127]
        self.assertEqual(len(non_ascii), 0, f"Found non-ASCII bytes in VietnameseNumber.bas: {non_ascii[:10]}")

    def test_ascii_safety_test_engine(self):
        """Verify test_engine.bas contains strictly 7-bit ASCII characters."""
        path = PROJECT_ROOT / "tests" / "test_engine.bas"
        self.assertTrue(path.is_file(), f"test_engine.bas not found at {path}")
        content = path.read_bytes()
        non_ascii = [b for b in content if b > 127]
        self.assertEqual(len(non_ascii), 0, f"Found non-ASCII bytes in test_engine.bas: {non_ascii[:10]}")

    def test_expected_cases_csv_validity(self):
        """Verify tests/expected_cases.csv contains all required columns and scenarios."""
        csv_path = PROJECT_ROOT / "tests" / "expected_cases.csv"
        self.assertTrue(csv_path.is_file(), f"expected_cases.csv not found at {csv_path}")

        with open(csv_path, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            required_cols = {
                "case_id",
                "api",
                "input",
                "expected",
                "zero_style",
                "thousand_style",
                "four_style",
                "decimal_mode",
                "expected_error",
            }
            self.assertTrue(
                required_cols.issubset(set(reader.fieldnames or [])),
                f"Missing columns in expected_cases.csv. Found: {reader.fieldnames}",
            )
            rows = list(reader)

        self.assertGreaterEqual(len(rows), 45, f"Expected at least 45 test cases, found {len(rows)}")

        # Verify presence of all key categories AC-01 through AC-12
        case_prefixes = {row["case_id"].split("-")[0] + "-" + row["case_id"].split("-")[1] for row in rows}
        for ac_num in range(1, 13):
            expected_prefix = f"AC-{ac_num:02d}"
            self.assertIn(
                expected_prefix,
                case_prefixes,
                f"Category {expected_prefix} not found in expected_cases.csv",
            )


class TestNumberEngineLiveCOM(unittest.TestCase):
    """Executes Vietnamese number engine via live Excel COM automation."""

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

        core_types = PROJECT_ROOT / "src" / "core" / "CoreTypes.bas"
        unicode_text = PROJECT_ROOT / "src" / "core" / "UnicodeText.bas"
        vn_number = PROJECT_ROOT / "src" / "core" / "VietnameseNumber.bas"
        test_engine = PROJECT_ROOT / "tests" / "test_engine.bas"

        cls.wb.VBProject.VBComponents.Import(str(core_types))
        cls.wb.VBProject.VBComponents.Import(str(unicode_text))
        cls.wb.VBProject.VBComponents.Import(str(vn_number))
        cls.wb.VBProject.VBComponents.Import(str(test_engine))

    @classmethod
    def tearDownClass(cls):
        if cls.wb:
            cls.wb.Close(SaveChanges=False)
        if cls.excel_session:
            cls.excel_session.__exit__(None, None, None)

    def test_in_memory_vba_runner(self):
        """Run the pure VBA in-memory test runner RunCoreTests()."""
        macro_name = f"'{self.wb_name}'!test_engine.RunCoreTests"
        result = self.excel.Run(macro_name)
        self.assertTrue(
            result.startswith("PASS:"),
            f"RunCoreTests() in VBA failed: {result}",
        )

    def test_all_expected_cases_matrix(self):
        """Execute all rows in tests/expected_cases.csv through Excel COM and assert results."""
        csv_path = PROJECT_ROOT / "tests" / "expected_cases.csv"
        with open(csv_path, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            rows = list(reader)

        run_macro = f"'{self.wb_name}'!test_engine.NumberToVietnameseWithOptions"
        try_macro = f"'{self.wb_name}'!test_engine.TryNumberToVietnameseWithOptions"
        err_msg_macro = f"'{self.wb_name}'!test_engine.TryConvertGetError"
        err_code_macro = f"'{self.wb_name}'!test_engine.GetNumberToVietnameseErrorCode"
        default_macro = f"'{self.wb_name}'!VietnameseNumber.NumberToVietnameseDefault"

        ERR_CODES = {
            "ERR_INVALID_NUMBER": -2147221504 + 2101,
            "ERR_OUT_OF_RANGE": -2147221504 + 2102,
            "ERR_INVALID_OPTIONS": -2147221504 + 2103,
        }

        for row in rows:
            case_id = row["case_id"]
            raw_input = row["input"]
            expected = row["expected"]
            zero_style = int(row["zero_style"])
            thousand_style = int(row["thousand_style"])
            four_style = int(row["four_style"])
            decimal_mode = int(row["decimal_mode"])
            expected_error = row["expected_error"].strip()

            # Parse input value
            if expected_error == "ERR_INVALID_NUMBER":
                input_val = raw_input  # non-numeric string like "abc"
            else:
                try:
                    if "." in raw_input:
                        input_val = float(raw_input)
                    else:
                        input_val = int(raw_input)
                except ValueError:
                    input_val = raw_input

            with self.subTest(case_id=case_id, input=raw_input):
                if expected_error:
                    # Verify TryNumberToVietnamese returns empty string
                    out_text = self.excel.Run(
                        try_macro,
                        input_val,
                        zero_style,
                        thousand_style,
                        four_style,
                        decimal_mode,
                    )
                    self.assertEqual(
                        out_text,
                        "",
                        f"[{case_id}] Expected error, but TryNumberToVietnamese returned '{out_text}'",
                    )
                    # Verify TryConvertGetError returns a friendly error message
                    err_msg = self.excel.Run(
                        err_msg_macro,
                        input_val,
                        zero_style,
                        thousand_style,
                        four_style,
                        decimal_mode,
                    )
                    self.assertTrue(
                        len(err_msg) > 0,
                        f"[{case_id}] Expected error message, but got empty string",
                    )
                    # Verify strict NumberToVietnamese raises the exact expected error code
                    err_code = self.excel.Run(
                        err_code_macro,
                        input_val,
                        zero_style,
                        thousand_style,
                        four_style,
                        decimal_mode,
                    )
                    expected_code = ERR_CODES.get(expected_error)
                    self.assertEqual(
                        err_code,
                        expected_code,
                        f"[{case_id}] Expected error code {expected_code} ({expected_error}), got {err_code}",
                    )
                else:
                    actual = self.excel.Run(
                        run_macro,
                        input_val,
                        zero_style,
                        thousand_style,
                        four_style,
                        decimal_mode,
                    )
                    self.assertEqual(
                        actual,
                        expected,
                        f"[{case_id}] Input={raw_input}: actual='{actual}', expected='{expected}'",
                    )

                    # For default options, also verify NumberToVietnameseDefault
                    if zero_style == 0 and thousand_style == 0 and four_style == 0 and decimal_mode == 0:
                        default_actual = self.excel.Run(default_macro, input_val)
                        self.assertEqual(
                            default_actual,
                            expected,
                            f"[{case_id}] NumberToVietnameseDefault mismatch: actual='{default_actual}', expected='{expected}'",
                        )

    def test_is_valid_number_api(self):
        """Verify IsValidNumber correctly validates numeric subtypes and rejects invalid inputs."""
        valid_macro = f"'{self.wb_name}'!test_engine.IsValidNumberBridge"

        # Valid numeric subtypes
        self.assertTrue(self.excel.Run(valid_macro, 0, ""))
        self.assertTrue(self.excel.Run(valid_macro, 125, ""))
        self.assertTrue(self.excel.Run(valid_macro, -150000, ""))
        self.assertTrue(self.excel.Run(valid_macro, 125.05, ""))
        self.assertTrue(self.excel.Run(valid_macro, 999999999999999, ""))
        self.assertTrue(self.excel.Run(valid_macro, -999999999999999, ""))

        # Invalid: out of range
        self.assertFalse(self.excel.Run(valid_macro, 1000000000000000, ""))
        self.assertFalse(self.excel.Run(valid_macro, -1000000000000000, ""))

        # Invalid: non-numeric strings
        self.assertFalse(self.excel.Run(valid_macro, "abc", ""))
        self.assertFalse(self.excel.Run(valid_macro, "", ""))


if __name__ == "__main__":
    unittest.main()
