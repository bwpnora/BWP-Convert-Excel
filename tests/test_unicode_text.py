"""
Unit tests for CoreTypes.bas and UnicodeText.bas
Verifies pure ASCII safety, core types declarations, exact Unicode code points for Vietnamese vocabulary,
complete casing tables covering all 134+ Vietnamese accented character pairs,
and casing functions on sample phrases both statically and via live Excel VBA COM execution.
"""

import re
import sys
import unittest
from pathlib import Path

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

# Expected 30 core vocabulary words
EXPECTED_VOCABULARY = {
    "WordKhong": "kh\u00f4ng",                                  # "không"
    "WordMot": "m\u1ed9t",                                      # "một"
    "WordHai": "hai",                                          # "hai"
    "WordBa": "ba",                                            # "ba"
    "WordBon": "b\u1ed1n",                                      # "bốn"
    "WordNam": "n\u0103m",                                      # "năm"
    "WordSau": "s\u00e1u",                                      # "sáu"
    "WordBay": "b\u1ea3y",                                      # "bảy"
    "WordTam": "t\u00e1m",                                      # "tám"
    "WordChin": "ch\u00edn",                                    # "chín"
    "WordMuoi": "m\u01b0\u1eddng".replace("ng", "i"),           # "mười" (\u01b0 \u1edd i)
    "WordMuoiChuc": "m\u01b0\u01a1i",                          # "mươi"
    "WordTram": "tr\u0103m",                                    # "trăm"
    "WordNghin": "ngh\u00ecn",                                  # "nghìn"
    "WordNgan": "ng\u00e0n",                                    # "ngàn"
    "WordTrieu": "tri\u1ec7u",                                  # "triệu"
    "WordTy": "t\u1ef7",                                        # "tỷ"
    "WordLe": "l\u1ebb",                                        # "lẻ"
    "WordLinh": "linh",                                        # "linh"
    "WordMotCuoi": "m\u1ed1t",                                  # "mốt"
    "WordTu": "t\u01b0",                                        # "tư"
    "WordLam": "l\u0103m",                                      # "lăm"
    "WordAm": "\u00c2m",                                        # "Âm"
    "WordPhay": "ph\u1ea9y",                                    # "phẩy"
    "WordDong": "\u0111\u1ed3ng",                               # "đồng"
    "WordDongHoa": "\u0110\u1ed3ng",                            # "Đồng"
    "WordChan": "ch\u1eb5n",                                    # "chẵn"
    "WordDolaMy": "\u0111\u00f4 la M\u1ef9",                    # "đô la Mỹ"
    "WordCent": "cent",                                        # "cent"
    "WordXu": "xu",                                            # "xu"
}

# Complete set of all 67 accented Vietnamese character pairs (134 characters)
VIETNAMESE_ACCENTED_PAIRS = [
    # a with diacritics (17 pairs)
    ("\u00e0", "\u00c0"),  # à, À
    ("\u00e1", "\u00c1"),  # á, Á
    ("\u1ea3", "\u1ea2"),  # ả, Ả
    ("\u00e3", "\u00c3"),  # ã, Ã
    ("\u1ea1", "\u1ea0"),  # ạ, Ạ
    ("\u0103", "\u0102"),  # ă, Ă
    ("\u1eb1", "\u1eb0"),  # ằ, Ằ
    ("\u1eaf", "\u1eae"),  # ắ, Ắ
    ("\u1eb3", "\u1eb2"),  # ẳ, Ẳ
    ("\u1eb5", "\u1eb4"),  # ẵ, Ẵ
    ("\u1eb7", "\u1eb6"),  # ặ, Ặ
    ("\u00e2", "\u00c2"),  # â, Â
    ("\u1ea7", "\u1ea6"),  # ầ, Ầ
    ("\u1ea5", "\u1ea4"),  # ấn, Ấ
    ("\u1ea9", "\u1ea8"),  # ẩ, Ẩ
    ("\u1eab", "\u1eaa"),  # ẫ, Ẫ
    ("\u1ead", "\u1eac"),  # ậ, Ậ
    # e with diacritics (11 pairs)
    ("\u00e8", "\u00c8"),  # è, È
    ("\u00e9", "\u00c9"),  # é, É
    ("\u1ebb", "\u1eba"),  # ẻ, Ẻ
    ("\u1ebd", "\u1ebc"),  # ẽ, Ẽ
    ("\u1eb9", "\u1eb8"),  # ẹ, Ẹ
    ("\u00ea", "\u00ca"),  # ê, Ê
    ("\u1ec1", "\u1ec0"),  # ề, Ề
    ("\u1ebf", "\u1ebe"),  # ế, Ế
    ("\u1ec3", "\u1ec2"),  # ể, Ể
    ("\u1ec5", "\u1ec4"),  # ễ, Ễ
    ("\u1ec7", "\u1ec6"),  # ệ, Ệ
    # i with diacritics (5 pairs)
    ("\u00ec", "\u00cc"),  # ì, Ì
    ("\u00ed", "\u00cd"),  # í, Í
    ("\u1ec9", "\u1ec8"),  # ỉ, Ỉ
    ("\u0129", "\u0128"),  # ĩ, Ĩ
    ("\u1ecb", "\u1eca"),  # ị, Ị
    # o with diacritics (17 pairs)
    ("\u00f2", "\u00d2"),  # ò, Ò
    ("\u00f3", "\u00d3"),  # ó, Ó
    ("\u1ecf", "\u1ece"),  # ỏ, Ỏ
    ("\u00f5", "\u00d5"),  # õ, Õ
    ("\u1ecd", "\u1ecc"),  # ọ, Ọ
    ("\u00f4", "\u00d4"),  # ô, Ô
    ("\u1ed3", "\u1ed2"),  # ồ, Ồ
    ("\u1ed1", "\u1ed0"),  # ố, Ố
    ("\u1ed5", "\u1ed4"),  # ổ, Ổ
    ("\u1ed7", "\u1ed6"),  # ỗ, Ỗ
    ("\u1ed9", "\u1ed8"),  # ộ, Ộ
    ("\u01a1", "\u01a0"),  # ơ, Ơ
    ("\u1edd", "\u1edc"),  # ờ, Ờ
    ("\u1edb", "\u1eda"),  # ớ, Ớ
    ("\u1edf", "\u1ede"),  # ở, Ở
    ("\u1ee1", "\u1ee0"),  # ỡ, Ỡ
    ("\u1ee3", "\u1ee2"),  # ợ, Ợ
    # u with diacritics (11 pairs)
    ("\u00f9", "\u00d9"),  # ù, Ù
    ("\u00fa", "\u00da"),  # ú, Ú
    ("\u1ee7", "\u1ee6"),  # ủ, Ủ
    ("\u0169", "\u0168"),  # ũ, Ũ
    ("\u1ee5", "\u1ee4"),  # ụ, Ụ
    ("\u01b0", "\u01af"),  # ư, Ư
    ("\u1eeb", "\u1eea"),  # ừ, Ừ
    ("\u1ee9", "\u1ee8"),  # ứ, Ứ
    ("\u1eed", "\u1eec"),  # ử, Ử
    ("\u1eef", "\u1eee"),  # ữ, Ữ
    ("\u1ef1", "\u1ef0"),  # ự, Ự
    # y with diacritics (5 pairs)
    ("\u1ef3", "\u1ef2"),  # ỳ, Ỳ
    ("\u00fd", "\u00dd"),  # ý, Ý
    ("\u1ef7", "\u1ef6"),  # ỷ, Ỷ
    ("\u1ef9", "\u1ef8"),  # ỹ, Ỹ
    ("\u1ef5", "\u1ef4"),  # ỵ, Ỵ
    # consonant d with stroke (1 pair)
    ("\u0111", "\u0110"),  # đ, Đ
]


def evaluate_vba_expr(expr: str) -> str:
    """Evaluate a VBA string expression composed of string literals and ChrW$(&H...)."""
    parts = re.split(r"\s+&\s+", expr.strip())
    res = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if part.startswith('"') and part.endswith('"'):
            res.append(part[1:-1])
        elif part.upper().startswith("CHRW$(") and part.endswith(")"):
            inner = part[6:-1].strip()
            if inner.upper().startswith("&H"):
                cp = int(inner[2:], 16)
            else:
                cp = int(inner)
            res.append(chr(cp))
        else:
            raise ValueError(f"Unknown VBA expression fragment: {part}")
    return "".join(res)


class TestCoreTypes(unittest.TestCase):
    """Verifies src/core/CoreTypes.bas contents and declarations."""

    @classmethod
    def setUpClass(cls):
        cls.core_types_path = PROJECT_ROOT / "src" / "core" / "CoreTypes.bas"
        cls.assertTrue(cls.core_types_path.is_file(), f"CoreTypes.bas not found at {cls.core_types_path}")
        cls.content = cls.core_types_path.read_text(encoding="ascii")

    def test_pure_ascii_safety(self):
        """Verify CoreTypes.bas contains strictly 7-bit ASCII characters."""
        non_ascii = [c for c in self.content if ord(c) > 127]
        self.assertEqual(len(non_ascii), 0, f"CoreTypes.bas contains non-ASCII characters: {non_ascii}")

    def test_constants_and_error_codes(self):
        """Verify MAX_SUPPORTED_VALUE and error codes are declared."""
        self.assertIn("Public Const MAX_SUPPORTED_VALUE As Double = 999999999999999#", self.content)
        self.assertIn("Public Const ERR_INVALID_NUMBER As Long = vbObjectError + 2101", self.content)
        self.assertIn("Public Const ERR_OUT_OF_RANGE As Long = vbObjectError + 2102", self.content)
        self.assertIn("Public Const ERR_INVALID_OPTIONS As Long = vbObjectError + 2103", self.content)

    def test_enums_declared(self):
        """Verify all required enums are present."""
        expected_enums = [
            "VnZeroStyle",
            "VnThousandStyle",
            "VnFourStyle",
            "VnDecimalMode",
            "VnCurrencyType",
            "VnCasingStyle",
        ]
        for enum_name in expected_enums:
            self.assertIn(f"Public Enum {enum_name}", self.content)

    def test_udts_and_decimal_places_field_name(self):
        """Verify UDTs are declared and field name is DecimalPlaces (not CustomDecimals!)."""
        self.assertIn("Public Type VnEngineOptions", self.content)
        self.assertIn("Public Type VnCurrencyOptions", self.content)
        self.assertIn("Public Type VnFormatOptions", self.content)
        # Check exact field name
        self.assertIn("DecimalPlaces As Integer", self.content)
        self.assertNotIn("CustomDecimals", self.content)

    def test_factory_functions(self):
        """Verify default factory functions are present."""
        self.assertIn("Public Function DefaultEngineOptions() As VnEngineOptions", self.content)
        self.assertIn("Public Function DefaultCurrencyOptions(", self.content)
        self.assertIn("Public Function DefaultFormatOptions() As VnFormatOptions", self.content)


class TestUnicodeTextEngine(unittest.TestCase):
    """Verifies src/core/UnicodeText.bas vocabulary, casing tables, and transformation functions."""

    @classmethod
    def setUpClass(cls):
        cls.unicode_path = PROJECT_ROOT / "src" / "core" / "UnicodeText.bas"
        cls.assertTrue(cls.unicode_path.is_file(), f"UnicodeText.bas not found at {cls.unicode_path}")
        cls.content = cls.unicode_path.read_text(encoding="ascii")

    def test_pure_ascii_safety(self):
        """Verify UnicodeText.bas contains strictly 7-bit ASCII characters."""
        non_ascii = [c for c in self.content if ord(c) > 127]
        self.assertEqual(len(non_ascii), 0, f"UnicodeText.bas contains non-ASCII characters: {non_ascii}")

    def test_all_vocabulary_code_points_match_vietnamese(self):
        """Verify all 30 vocabulary functions in UnicodeText.bas evaluate to exact Vietnamese words."""
        func_pattern = re.compile(
            r"Public Function (Word[A-Za-z0-9_]+)\(\) As String\s+(.*?)\s+End Function",
            re.DOTALL,
        )
        matches = func_pattern.findall(self.content)
        extracted_funcs = {}
        for func_name, body in matches:
            assign_pattern = re.compile(rf"{func_name}\s*=\s*(.+)")
            assign_match = assign_pattern.search(body)
            self.assertIsNotNone(assign_match, f"Could not find assignment in function {func_name}")
            expr = assign_match.group(1).strip()
            val = evaluate_vba_expr(expr)
            extracted_funcs[func_name] = val

        self.assertEqual(len(extracted_funcs), len(EXPECTED_VOCABULARY))
        for func_name, expected_str in EXPECTED_VOCABULARY.items():
            self.assertIn(func_name, extracted_funcs, f"Missing function {func_name} in UnicodeText.bas")
            actual_str = extracted_funcs[func_name]
            self.assertEqual(
                actual_str,
                expected_str,
                f"Mismatch in {func_name}: actual={ascii(actual_str)}, expected={ascii(expected_str)}",
            )
            # Guarantee zero corrupted characters
            self.assertNotIn("?", actual_str)

    def test_casing_table_coverage(self):
        """Verify casing tables in UnicodeText.bas cover all 67 accented pairs (134 characters)."""
        pair_pattern = re.compile(r"MapPair\s+&H([0-9A-Fa-f]+),\s+&H([0-9A-Fa-f]+)")
        matches = pair_pattern.findall(self.content)
        mapped_pairs = [(int(low, 16), int(up, 16)) for low, up in matches]

        # Verify all 67 accented pairs are included
        self.assertGreaterEqual(len(mapped_pairs), len(VIETNAMESE_ACCENTED_PAIRS))

        mapped_dict_up = dict(mapped_pairs)
        mapped_dict_low = {up: low for low, up in mapped_pairs}

        for low_char, up_char in VIETNAMESE_ACCENTED_PAIRS:
            low_cp = ord(low_char)
            up_cp = ord(up_char)
            self.assertIn(
                low_cp,
                mapped_dict_up,
                f"Missing lowercase code point U+{low_cp:04X} ({ascii(low_char)}) in MapPair",
            )
            self.assertEqual(
                mapped_dict_up[low_cp],
                up_cp,
                f"Lowercase U+{low_cp:04X} did not map to uppercase U+{up_cp:04X}",
            )
            self.assertIn(
                up_cp,
                mapped_dict_low,
                f"Missing uppercase code point U+{up_cp:04X} ({ascii(up_char)}) in MapPair",
            )
            self.assertEqual(
                mapped_dict_low[up_cp],
                low_cp,
                f"Uppercase U+{up_cp:04X} did not map to lowercase U+{low_cp:04X}",
            )

    def test_casing_logic_on_sample_phrases(self):
        """Verify casing conversion logic on sample phrases: dong, do la My, and full phrase."""
        # Build map from UnicodeText.bas pairs
        pair_pattern = re.compile(r"MapPair\s+&H([0-9A-Fa-f]+),\s+&H([0-9A-Fa-f]+)")
        matches = pair_pattern.findall(self.content)
        upper_map = {int(low, 16): int(up, 16) for low, up in matches}
        lower_map = {int(up, 16): int(low, 16) for low, up in matches}
        # Include ASCII Latin letter conversions (a-z <-> A-Z)
        for i in range(97, 123):
            upper_map[i] = i - 32
            lower_map[i - 32] = i

        def to_upper(text: str) -> str:
            res = []
            for ch in text:
                cp = ord(ch)
                res.append(chr(upper_map.get(cp, cp)))
            return "".join(res)

        def to_lower(text: str) -> str:
            res = []
            for ch in text:
                cp = ord(ch)
                res.append(chr(lower_map.get(cp, cp)))
            return "".join(res)

        def cap_first(text: str) -> str:
            for idx, ch in enumerate(text):
                if not ch.isspace():
                    cp = ord(ch)
                    new_ch = chr(upper_map.get(cp, cp))
                    return text[:idx] + new_ch + text[idx + 1 :]
            return text

        # Sample 1: "đồng"
        s1 = "\u0111\u1ed3ng"
        self.assertEqual(cap_first(s1), "\u0110\u1ed3ng")
        self.assertEqual(to_upper(s1), "\u0110\u1ed2NG")
        self.assertEqual(to_lower(cap_first(s1)), s1)

        # Sample 2: "đô la Mỹ"
        s2 = "\u0111\u00f4 la M\u1ef9"
        self.assertEqual(cap_first(s2), "\u0110\u00f4 la M\u1ef9")
        self.assertEqual(to_upper(s2), "\u0110\u00d4 LA M\u1ef8")
        self.assertEqual(to_lower("\u0110\u00d4 LA M\u1ef8"), "\u0111\u00f4 la m\u1ef9")

        # Sample 3: "một triệu không trăm lẻ một nghìn đồng chẵn"
        s3 = (
            "m\u1ed9t tri\u1ec7u kh\u00f4ng tr\u0103m l\u1ebb "
            "m\u1ed9t ngh\u00ecn \u0111\u1ed3ng ch\u1eb5n"
        )
        s3_cap = (
            "M\u1ed9t tri\u1ec7u kh\u00f4ng tr\u0103m l\u1ebb "
            "m\u1ed9t ngh\u00ecn \u0111\u1ed3ng ch\u1eb5n"
        )
        s3_upper = (
            "M\u1ed8T TRI\u1ec6U KH\u00d4NG TR\u0102M L\u1eba "
            "M\u1ed8T NGH\u00ccN \u0110\u1ed2NG CH\u1eb4N"
        )
        self.assertEqual(cap_first(s3), s3_cap)
        self.assertEqual(to_upper(s3), s3_upper)
        self.assertEqual(to_lower(s3_upper), s3)

        # Edge cases
        self.assertEqual(cap_first(""), "")
        self.assertEqual(cap_first("   " + s1), "   " + "\u0110\u1ed3ng")
        self.assertEqual(to_upper("123 abc!"), "123 ABC!")
        self.assertEqual(to_lower("123 ABC!"), "123 abc!")

    def test_live_excel_vba_execution(self):
        """Verify actual VBA compilation and execution in an isolated Excel COM session."""
        from scripts.build import ExcelSession

        try:
            session_ctx = ExcelSession(visible=False, display_alerts=False)
            session = session_ctx.__enter__()
        except Exception as e:
            self.skipTest(f"Excel COM not available on this host: {e}")
            return

        core_types_file = PROJECT_ROOT / "src" / "core" / "CoreTypes.bas"
        unicode_text_file = PROJECT_ROOT / "src" / "core" / "UnicodeText.bas"

        try:
            wb = session.excel.Workbooks.Add()
            try:
                wb.VBProject.VBComponents.Import(str(core_types_file))
                wb.VBProject.VBComponents.Import(str(unicode_text_file))

                # Test vocabulary functions in live Excel VBA
                self.assertEqual(session.excel.Run("WordDong"), "\u0111\u1ed3ng")
                self.assertEqual(session.excel.Run("WordDongHoa"), "\u0110\u1ed3ng")
                self.assertEqual(session.excel.Run("WordDolaMy"), "\u0111\u00f4 la M\u1ef9")
                self.assertEqual(session.excel.Run("WordChan"), "ch\u1eb5n")
                self.assertEqual(session.excel.Run("WordNghin"), "ngh\u00ecn")
                self.assertEqual(session.excel.Run("WordTy"), "t\u1ef7")
                self.assertEqual(session.excel.Run("WordLe"), "l\u1ebb")
                self.assertEqual(session.excel.Run("WordLinh"), "linh")

                # Test casing functions in live Excel VBA
                # Sample 1: "đồng"
                self.assertEqual(
                    session.excel.Run("CapitalizeFirst", "\u0111\u1ed3ng"),
                    "\u0110\u1ed3ng",
                )
                self.assertEqual(
                    session.excel.Run("ToUnicodeUpper", "\u0111\u1ed3ng"),
                    "\u0110\u1ed2NG",
                )
                self.assertEqual(
                    session.excel.Run("ToUnicodeLower", "\u0110\u1ed2NG"),
                    "\u0111\u1ed3ng",
                )

                # Sample 2: "đô la Mỹ"
                self.assertEqual(
                    session.excel.Run("CapitalizeFirst", "\u0111\u00f4 la M\u1ef9"),
                    "\u0110\u00f4 la M\u1ef9",
                )
                self.assertEqual(
                    session.excel.Run("ToUnicodeUpper", "\u0111\u00f4 la M\u1ef9"),
                    "\u0110\u00d4 LA M\u1ef8",
                )
                self.assertEqual(
                    session.excel.Run("ToUnicodeLower", "\u0110\u00d4 LA M\u1ef8"),
                    "\u0111\u00f4 la m\u1ef9",
                )

                # Sample 3: "một triệu không trăm lẻ một nghìn đồng chẵn"
                s3 = (
                    "m\u1ed9t tri\u1ec7u kh\u00f4ng tr\u0103m l\u1ebb "
                    "m\u1ed9t ngh\u00ecn \u0111\u1ed3ng ch\u1eb5n"
                )
                s3_cap = (
                    "M\u1ed9t tri\u1ec7u kh\u00f4ng tr\u0103m l\u1ebb "
                    "m\u1ed9t ngh\u00ecn \u0111\u1ed3ng ch\u1eb5n"
                )
                s3_upper = (
                    "M\u1ed8T TRI\u1ec6U KH\u00d4NG TR\u0102M L\u1eba "
                    "M\u1ed8T NGH\u00ccN \u0110\u1ed2NG CH\u1eb4N"
                )
                self.assertEqual(session.excel.Run("CapitalizeFirst", s3), s3_cap)
                self.assertEqual(session.excel.Run("ToUnicodeUpper", s3), s3_upper)
                self.assertEqual(session.excel.Run("ToUnicodeLower", s3_upper), s3)

                # Edge cases in live Excel VBA
                self.assertEqual(session.excel.Run("CapitalizeFirst", ""), "")
                self.assertEqual(
                    session.excel.Run("CapitalizeFirst", "   \u0111\u1ed3ng"),
                    "   \u0110\u1ed3ng",
                )
            finally:
                wb.Close(SaveChanges=False)
        finally:
            session_ctx.__exit__(None, None, None)


if __name__ == "__main__":
    unittest.main()
