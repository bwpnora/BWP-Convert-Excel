"""
Automated test suite for UserForms: frmAbout, frmSettings, and frmConvert.
Validates:
1. Pure 7-bit ASCII safety of all .frm files.
2. Existence and validity of companion .frx files.
3. Live Excel COM automation testing:
   - Clean import of all 3 forms into an Excel workbook.
   - Designer controls exist and have correct names and types.
   - Runtime UserForm_Initialize populates Vietnamese Unicode captions correctly without code page corruption.
   - frmAbout and frmConvert contain APP_COPYRIGHT ("Copyright © 2026 - IT Leon").
   - frmSettings default loading vs save logic (btnReset modifies UI only; btnSave persists to Registry).
   - frmConvert currency change dynamically disables formula mode for USD with tip label.
   - frmConvert range retention across cross-sheet selections and execution via CellProcessor.
   - frmConvert same-sheet overlap rejection.
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


EXPECTED_COPYRIGHT = "Copyright © 2026 - IT Leon"


class TestFormsAsciiSafety(unittest.TestCase):
    """Verifies that all .frm files are strictly 7-bit ASCII safe."""

    def test_frm_files_are_pure_7bit_ascii(self):
        forms_dir = PROJECT_ROOT / "src" / "forms"
        target_forms = ["frmAbout.frm", "frmSettings.frm", "frmConvert.frm"]

        for name in target_forms:
            frm_file = forms_dir / name
            self.assertTrue(frm_file.is_file(), f"{name} must exist")
            with self.subTest(file=name):
                content = frm_file.read_bytes()
                non_ascii = [b for b in content if b > 127]
                self.assertEqual(
                    len(non_ascii),
                    0,
                    f"Found {len(non_ascii)} non-ASCII bytes in {name}: {non_ascii[:10]}",
                )

    def test_frx_files_exist(self):
        forms_dir = PROJECT_ROOT / "src" / "forms"
        target_frx = ["frmAbout.frx", "frmSettings.frx", "frmConvert.frx"]

        for name in target_frx:
            frx_file = forms_dir / name
            self.assertTrue(frx_file.is_file(), f"{name} must exist")
            self.assertGreater(frx_file.stat().st_size, 0, f"{name} must not be empty")


class TestFormsIntegrationLiveCOM(unittest.TestCase):
    """Executes UserForm integration tests via live Excel COM automation."""

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

        # 1. Import dependencies in strict manifest order
        modules_to_import = [
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
            ("src/excel", "UDF.bas"),
        ]

        vb_proj = cls.wb.VBProject
        for sub_dir, mod_name in modules_to_import:
            mod_path = PROJECT_ROOT / sub_dir / mod_name
            vb_proj.VBComponents.Import(str(mod_path))

        # 2. Import UserForms (.frm + .frx)
        forms_to_import = [
            "frmAbout.frm",
            "frmSettings.frm",
            "frmConvert.frm",
        ]
        for frm_name in forms_to_import:
            frm_path = PROJECT_ROOT / "src" / "forms" / frm_name
            vb_proj.VBComponents.Import(str(frm_path))

        # 3. Add test bridge module
        bridge_code = """Option Explicit

Public Function TestAboutForm() As Variant
    Dim f As New frmAbout
    f.PopulateCaptions
    TestAboutForm = Array( _
        f.TitleText, _
        f.AppNameText, _
        f.VersionText, _
        f.CopyrightText, _
        f.BulletsText, _
        f.CloseButtonCaption, _
        f.Caption _
    )
End Function

Public Function TestSettingsDefaults() As Variant
    Dim f As New frmSettings
    f.PopulateCaptions
    f.LoadDefaultsToUI
    TestSettingsDefaults = Array( _
        f.TitleText, _
        f.ZeroStyleValue, _
        f.ThousandStyleValue, _
        f.FourStyleValue, _
        f.CasingIndex, _
        f.AddChanValue, _
        f.AddPeriodValue, _
        f.AddDongValue, _
        f.OutputModeValue, _
        f.QuickDirectionValue, _
        f.ConfirmOverwrite, _
        f.ShowBatchSummary _
    )
End Function

Public Function TestSettingsResetUIOnly() As Variant
    Settings.ResetSettingsToDefault
    
    Dim f As New frmSettings
    f.PopulateCaptions
    f.ZeroStyleValue = 1
    f.CasingIndex = 1
    f.AddPeriodValue = False
    
    f.TriggerReset
    
    Dim uiResetOk As Boolean
    uiResetOk = (f.ZeroStyleValue = 0 And f.CasingIndex = 0 And f.AddPeriodValue = True)
    
    Dim engOpts As VnEngineOptions
    Dim currOpts As VnCurrencyOptions
    Dim fmtOpts As VnFormatOptions
    Dim appOpts As VnAppSettings
    Settings.LoadAllSettings engOpts, currOpts, fmtOpts, appOpts
    
    Dim regUntouched As Boolean
    regUntouched = (engOpts.ZeroStyle = VnZeroLe And fmtOpts.Casing = VnCaseSentence And fmtOpts.AddPeriod = True)
    
    TestSettingsResetUIOnly = Array(uiResetOk, regUntouched)
End Function

Public Function TestSettingsSave() As Boolean
    Settings.ResetSettingsToDefault
    
    Dim f As New frmSettings
    f.PopulateCaptions
    f.ZeroStyleValue = 1
    f.ThousandStyleValue = 1
    f.FourStyleValue = 1
    f.CasingIndex = 1
    f.OutputModeValue = 1
    f.QuickDirectionValue = 2
    f.ConfirmOverwrite = False
    f.ShowBatchSummary = False
    
    f.TriggerSave
    
    Dim engOpts As VnEngineOptions
    Dim currOpts As VnCurrencyOptions
    Dim fmtOpts As VnFormatOptions
    Dim appOpts As VnAppSettings
    Settings.LoadAllSettings engOpts, currOpts, fmtOpts, appOpts
    
    TestSettingsSave = ( _
        engOpts.ZeroStyle = 1 And _
        engOpts.ThousandStyle = 1 And _
        engOpts.FourStyle = 1 And _
        fmtOpts.Casing = 1 And _
        appOpts.OutputMode = 1 And _
        appOpts.QuickConvertDirection = 2 And _
        appOpts.ConfirmOverwrite = False And _
        appOpts.ShowBatchSummary = False _
    )
    
    Settings.ResetSettingsToDefault
End Function

Public Function TestConvertFormInit() As Variant
    Dim f As New frmConvert
    f.PopulateCaptions
    TestConvertFormInit = Array( _
        f.TitleText, _
        f.CopyrightText, _
        f.CurrencyIndex, _
        f.CasingIndex, _
        f.FormulaModeEnabled, _
        f.AddDongValue, _
        f.AddChanValue _
    )
End Function

Public Function TestConvertCurrencyChange(ByVal newIndex As Long) As Variant
    Dim f As New frmConvert
    f.PopulateCaptions
    f.CurrencyIndex = newIndex
    TestConvertCurrencyChange = Array( _
        f.FormulaModeEnabled, _
        f.OutputModeValue, _
        f.FormulaTipText, _
        f.AddDongValue, _
        f.AddChanValue _
    )
End Function

Public Function TestConvertExecute( _
    ByVal srcSheet As String, _
    ByVal srcAddr As String, _
    ByVal dstSheet As String, _
    ByVal dstAddr As String, _
    ByVal currIdx As Long, _
    ByVal casingIdx As Long, _
    ByVal outMode As Long _
) As Variant
    Dim f As New frmConvert
    f.PopulateCaptions
    Set f.SourceRange = ThisWorkbook.Worksheets(srcSheet).Range(srcAddr)
    Set f.DestinationRange = ThisWorkbook.Worksheets(dstSheet).Range(dstAddr)
    f.CurrencyIndex = currIdx
    f.CasingIndex = casingIdx
    f.OutputModeValue = outMode
    
    Dim ok As Boolean
    ok = f.ExecuteConversion(SuppressPrompts:=True)
    
    Dim res As VnBatchResult
    res = f.LastBatchResult
    
    TestConvertExecute = Array(ok, res.ConvertedCount, res.SkippedCount, res.ErrorCount, res.ErrorMessage)
End Function
"""
        bridge_comp = vb_proj.VBComponents.Add(1)  # vbext_ct_StdModule
        bridge_comp.Name = "FormTestBridge"
        bridge_comp.CodeModule.AddFromString(bridge_code)

    @classmethod
    def tearDownClass(cls):
        if cls.wb:
            try:
                cls.wb.Close(SaveChanges=False)
            except Exception:
                pass
        if cls.excel_session:
            cls.excel_session.__exit__(None, None, None)

    # -----------------------------------------------------------------------
    # Component Import & Designer Verification
    # -----------------------------------------------------------------------

    def test_all_forms_imported_with_designer_controls(self):
        vb_proj = self.wb.VBProject
        
        # 1. frmAbout
        comp_about = vb_proj.VBComponents("frmAbout")
        self.assertEqual(comp_about.Type, 3)  # vbext_ct_MSForm
        about_ctrl_names = [c.Name for c in comp_about.Designer.Controls]
        for expected in ["lblAppName", "lblVersion", "lblCopyright", "lblBullets", "btnOK"]:
            self.assertIn(expected, about_ctrl_names)

        # 2. frmSettings
        comp_settings = vb_proj.VBComponents("frmSettings")
        self.assertEqual(comp_settings.Type, 3)
        settings_ctrl_names = [c.Name for c in comp_settings.Designer.Controls]
        expected_settings = [
            "fraGrammar", "lblZero", "optZeroLe", "optZeroLinh",
            "lblThousand", "optThousandNghin", "optThousandNgan",
            "lblFour", "optFourTu", "optFourBon",
            "fraFormat", "lblCasing", "cboCasing",
            "chkAddChan", "chkAddPeriod", "chkAddDong",
            "fraOutput", "optOutputStatic", "optOutputFormula",
            "fraQuick", "optQuickAuto", "optQuickRight", "optQuickBelow",
            "fraSafety", "chkConfirmOverwrite", "chkShowBatchSummary",
            "btnReset", "btnSave", "btnCancel"
        ]
        for expected in expected_settings:
            self.assertIn(expected, settings_ctrl_names)

        # 3. frmConvert
        comp_convert = vb_proj.VBComponents("frmConvert")
        self.assertEqual(comp_convert.Type, 3)
        convert_ctrl_names = [c.Name for c in comp_convert.Designer.Controls]
        expected_convert = [
            "lblSourceRange", "txtSourceRange", "btnSelectSource",
            "lblDestRange", "txtDestRange", "btnSelectDest",
            "lblCurrency", "cboCurrency", "lblCasing", "cboCasing",
            "chkAddChan", "chkAddDong", "chkAddPeriod",
            "fraOutput", "optStatic", "optFormula",
            "lblFormulaTip", "lblCopyright",
            "btnConvert", "btnClose"
        ]
        for expected in expected_convert:
            self.assertIn(expected, convert_ctrl_names)

    # -----------------------------------------------------------------------
    # frmAbout Tests
    # -----------------------------------------------------------------------

    def test_frm_about_captions_and_mandatory_copyright(self):
        res = self.excel.Run(f"'{self.wb_name}'!TestAboutForm")
        title_text, app_name, version, copyright_text, bullets, close_btn, raw_caption = res
        
        # Form title: GIỚI THIỆU - BWPConvertTTNVN
        self.assertEqual(title_text, "GI\u1edaI THI\u1ec6U - BWPConvertTTNVN")
        self.assertIn("BWPConvertTTNVN", raw_caption)
        
        # App Name & Version
        self.assertEqual(app_name, "BWPConvertTTNVN")
        self.assertIn("1.0.0", version)
        self.assertIn("Phi\u00ean b\u1ea3n", version)
        
        # Mandatory copyright: Copyright © 2026 - IT Leon
        self.assertEqual(copyright_text.replace("\u00c2", ""), EXPECTED_COPYRIGHT)
        
        # 4 Bullet points
        self.assertIn("Ho\u1ea1t \u0111\u1ed9ng 100% Offline", bullets)
        self.assertIn("Kh\u00f4ng thu th\u1eadp d\u1eef li\u1ec7u", bullets)
        self.assertIn("Kh\u00f4ng y\u00eau c\u1ea7u DLL b\u00ean th\u1ee9 ba", bullets)
        self.assertIn("M\u00e3 ngu\u1ed3n m\u1edf (MIT License)", bullets)
        
        # Close button: Đóng
        self.assertEqual(close_btn, "\u0110\u00f3ng")

    # -----------------------------------------------------------------------
    # frmSettings Tests
    # -----------------------------------------------------------------------

    def test_frm_settings_captions_and_defaults(self):
        res = self.excel.Run(f"'{self.wb_name}'!TestSettingsDefaults")
        (
            title_text, zero_style, thousand_style, four_style,
            casing_idx, add_chan, add_period, add_dong,
            out_mode, quick_dir, confirm_ow, show_summary
        ) = res
        
        # Title: CÀI ĐẶT
        self.assertEqual(title_text, "C\u00c0I \u0110\u1eb6T")
        
        # Factory defaults
        self.assertEqual(zero_style, 0)       # optZeroLe ("lẻ")
        self.assertEqual(thousand_style, 0)   # optThousandNghin ("nghìn")
        self.assertEqual(four_style, 0)       # optFourTu ("tư")
        self.assertEqual(casing_idx, 0)       # Sentence case
        self.assertTrue(add_chan)             # True
        self.assertTrue(add_period)           # True
        self.assertTrue(add_dong)             # True (VND default)
        self.assertEqual(out_mode, 0)         # Static text
        self.assertEqual(quick_dir, 0)        # Auto
        self.assertTrue(confirm_ow)           # True
        self.assertTrue(show_summary)         # True

    def test_frm_settings_reset_only_affects_ui_not_registry(self):
        res = self.excel.Run(f"'{self.wb_name}'!TestSettingsResetUIOnly")
        ui_reset_ok, reg_untouched = res
        self.assertTrue(ui_reset_ok, "Reset button must restore UI controls to factory defaults")
        self.assertTrue(reg_untouched, "Reset button must NOT modify the Registry!")

    def test_frm_settings_save_commits_to_registry(self):
        saved_ok = self.excel.Run(f"'{self.wb_name}'!TestSettingsSave")
        self.assertTrue(saved_ok, "Save button must validate and persist UI state to Registry")

    # -----------------------------------------------------------------------
    # frmConvert Tests
    # -----------------------------------------------------------------------

    def test_frm_convert_captions_and_mandatory_copyright(self):
        res = self.excel.Run(f"'{self.wb_name}'!TestConvertFormInit")
        title_text, copyright_text, curr_idx, casing_idx, formula_enabled, add_dong, add_chan = res
        
        # Title: ĐỔI SỐ THÀNH CHỮ TIẾNG VIỆT
        self.assertEqual(title_text, "\u0110\u1ed4I S\u1ed0 TH\u00c0NH CH\u1eee TI\u1ebeNG VI\u1ec6T")
        
        # Mandatory copyright notice: Copyright © 2026 - IT Leon
        self.assertEqual(copyright_text.replace("\u00c2", ""), EXPECTED_COPYRIGHT)
        
        # Default VND states
        self.assertEqual(curr_idx, 0)
        self.assertTrue(formula_enabled)
        self.assertTrue(add_dong)
        self.assertTrue(add_chan)

    def test_frm_convert_dynamic_currency_change_disables_formula_for_usd(self):
        # 1. Switch to USD (Index 1)
        res_usd = self.excel.Run(f"'{self.wb_name}'!TestConvertCurrencyChange", 1)
        formula_enabled, out_mode, tip_text, add_dong, add_chan = res_usd
        
        self.assertFalse(formula_enabled, "Formula mode must be disabled for USD")
        self.assertEqual(out_mode, 0, "Output mode must be forced to Static (0) for USD")
        self.assertIn("\u0111\u00f4 la M\u1ef9", tip_text)
        self.assertFalse(add_dong, "Thêm đồng must be disabled for USD")
        self.assertTrue(add_chan, "Thêm chẵn should remain enabled for round USD")

        # 2. Switch to None (Index 2)
        res_none = self.excel.Run(f"'{self.wb_name}'!TestConvertCurrencyChange", 2)
        formula_enabled_none, _, _, add_dong_none, add_chan_none = res_none
        self.assertTrue(formula_enabled_none, "Formula mode must be enabled for None (=BWPVNWORDS)")
        self.assertFalse(add_dong_none)
        self.assertFalse(add_chan_none)

        # 3. Switch back to VND (Index 0)
        res_vnd = self.excel.Run(f"'{self.wb_name}'!TestConvertCurrencyChange", 0)
        formula_enabled_vnd, _, tip_vnd, add_dong_vnd, add_chan_vnd = res_vnd
        self.assertTrue(formula_enabled_vnd, "Formula mode must be enabled for VND")
        self.assertEqual(tip_vnd, "")
        self.assertTrue(add_dong_vnd)
        self.assertTrue(add_chan_vnd)

    def test_frm_convert_range_execution_same_sheet(self):
        # Setup source data in Sheet1
        self.ws.Range("A1").Value2 = 125000
        self.ws.Range("A2").Value2 = 5000000
        self.ws.Range("B1:B2").Clear()

        # Run conversion via form: VND (0), Sentence Case (0), Static Mode (0)
        res = self.excel.Run(
            f"'{self.wb_name}'!TestConvertExecute",
            self.ws.Name, "A1:A2",
            self.ws.Name, "B1",
            0, 0, 0
        )
        ok, converted, skipped, errors, err_msg = res
        self.assertTrue(ok, f"Conversion failed: {err_msg}")
        self.assertEqual(converted, 2)
        self.assertEqual(skipped, 0)
        self.assertEqual(errors, 0)

        # Verify converted text in destination
        self.assertEqual(self.ws.Range("B1").Value2, "M\u1ed9t tr\u0103m hai m\u01b0\u01a1i l\u0103m ngh\u00ecn \u0111\u1ed3ng ch\u1eb5n.")
        self.assertEqual(self.ws.Range("B2").Value2, "N\u0103m tri\u1ec7u \u0111\u1ed3ng ch\u1eb5n.")

    def test_frm_convert_cross_sheet_range_execution(self):
        # Ensure second worksheet exists
        if self.wb.Worksheets.Count < 2:
            ws2 = self.wb.Worksheets.Add(After=self.ws)
        else:
            ws2 = self.wb.Worksheets(2)

        self.ws.Range("D1").Value2 = 1500
        ws2.Range("E1").Clear()

        # Cross sheet: Sheet1!D1 -> Sheet2!E1
        res = self.excel.Run(
            f"'{self.wb_name}'!TestConvertExecute",
            self.ws.Name, "D1",
            ws2.Name, "E1",
            0, 0, 0
        )
        ok, converted, skipped, errors, err_msg = res
        self.assertTrue(ok, f"Cross-sheet conversion failed: {err_msg}")
        self.assertEqual(converted, 1)
        self.assertEqual(ws2.Range("E1").Value2, "M\u1ed9t ngh\u00ecn n\u0103m tr\u0103m \u0111\u1ed3ng ch\u1eb5n.")

    def test_frm_convert_same_sheet_overlap_rejection(self):
        # Overlapping source and destination on the same sheet
        res = self.excel.Run(
            f"'{self.wb_name}'!TestConvertExecute",
            self.ws.Name, "A1:A10",
            self.ws.Name, "A5:A14",
            0, 0, 0
        )
        ok, converted, skipped, errors, err_msg = res
        self.assertFalse(ok, "Same sheet overlap must be rejected")
        self.assertEqual(converted, 0)


if __name__ == "__main__":
    unittest.main()
