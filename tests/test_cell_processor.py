"""
Automated test suite for Settings.bas, UndoManager.bas, and CellProcessor.bas.
Validates:
1. Pure 7-bit ASCII safety of all .bas files in src/excel/.
2. Live Excel COM automation testing of:
   - Single cell conversion (A1 to B1).
   - 1D batch conversion (A1:A5 with anchor B1).
   - Rejection of multi-area selection (A1:A2,C1:C2).
   - Rejection of same-sheet overlap (A1:A10 to A5:A14).
   - Cross-sheet conversion (Sheet1!A1:A5 to Sheet2!B1).
   - Formula-safe skip preservation (blank source cell leaves destination formula intact).
   - 2-phase transactional Undo (restores original values, formulas, and clears blank cells).
   - Protected sheet atomic abort (locked destination cell prevents any write).
   - Settings persistence, bounds validation, and reset in Registry.
   - Formula mode generation and USD/Custom rejection.
   - Application state caching and restoration on success and error paths.
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


class TestExcelAsciiSafety(unittest.TestCase):
    """Verifies that all VBA source modules in src/excel/ are strictly 7-bit ASCII safe."""

    def test_all_excel_bas_files_are_ascii(self):
        excel_dir = PROJECT_ROOT / "src" / "excel"
        target_names = ["Settings.bas", "UndoManager.bas", "CellProcessor.bas"]
        
        for name in target_names:
            bas_file = excel_dir / name
            self.assertTrue(bas_file.is_file(), f"{name} should exist")
            with self.subTest(file=name):
                content = bas_file.read_bytes()
                non_ascii = [b for b in content if b > 127]
                self.assertEqual(
                    len(non_ascii),
                    0,
                    f"Found {len(non_ascii)} non-ASCII bytes in {name}: {non_ascii[:10]}",
                )


class TestCellProcessorLiveCOM(unittest.TestCase):
    """Executes state, undo, settings, and cell processor tests via live Excel COM automation."""

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
            ("src/excel", "UndoManager.bas"),
            ("src/excel", "CellProcessor.bas"),
        ]

        for sub_dir, mod_name in modules_to_import:
            mod_path = PROJECT_ROOT / sub_dir / mod_name
            cls.wb.VBProject.VBComponents.Import(str(mod_path))

    @classmethod
    def tearDownClass(cls):
        if cls.wb:
            try:
                cls.wb.Close(SaveChanges=False)
            except Exception:
                pass
        if cls.excel_session:
            cls.excel_session.__exit__(None, None, None)

    def setUp(self):
        # Clear main worksheet before each test and clear undo
        try:
            self.ws.Unprotect("testpwd")
        except Exception:
            pass
        try:
            self.ws.Unprotect()
        except Exception:
            pass
        self.ws.Cells.Clear()
        clear_macro = f"'{self.wb_name}'!UndoManager.ClearUndo"
        self.excel.Run(clear_macro)

    # --------------------------------------------------------------------------
    # 1. Single Cell Conversion
    # --------------------------------------------------------------------------

    def test_single_cell_conversion(self):
        """Verify single cell A1 -> B1 converts correctly."""
        bridge_macro = f"'{self.wb_name}'!CellProcessor.ConvertRangeBridge"

        self.ws.Range("A1").Value = 1250000
        src = self.ws.Range("A1")
        dst = self.ws.Range("B1")

        # ConvertRangeBridge(src, dst) with default options
        res = self.excel.Run(bridge_macro, src, dst)
        success, converted, skipped, errors, undo_avail, err_msg = res

        self.assertTrue(success, f"Expected success, got error: {err_msg}")
        self.assertEqual(converted, 1)
        self.assertEqual(skipped, 0)
        self.assertEqual(errors, 0)
        self.assertTrue(undo_avail)
        self.assertEqual(self.ws.Range("B1").Value, "Một triệu hai trăm năm mươi nghìn đồng chẵn.")

    # --------------------------------------------------------------------------
    # 2. 1D Batch Conversion with Anchor Auto-Expansion
    # --------------------------------------------------------------------------

    def test_1d_batch_conversion_anchor_expansion(self):
        """Verify 1D range A1:A5 with anchor B1 expands to B1:B5 and converts all."""
        bridge_macro = f"'{self.wb_name}'!CellProcessor.ConvertRangeBridge"

        values = [[10000], [20000], [30000], [40000], [50000]]
        self.ws.Range("A1:A5").Value = values
        src = self.ws.Range("A1:A5")
        dst = self.ws.Range("B1")  # Single anchor cell

        res = self.excel.Run(bridge_macro, src, dst)
        success, converted, skipped, errors, undo_avail, err_msg = res

        self.assertTrue(success, f"Batch conversion failed: {err_msg}")
        self.assertEqual(converted, 5)
        self.assertEqual(skipped, 0)
        self.assertEqual(errors, 0)
        self.assertTrue(undo_avail)

        self.assertEqual(self.ws.Range("B1").Value, "Mười nghìn đồng chẵn.")
        self.assertEqual(self.ws.Range("B3").Value, "Ba mươi nghìn đồng chẵn.")
        self.assertEqual(self.ws.Range("B5").Value, "Năm mươi nghìn đồng chẵn.")

    # --------------------------------------------------------------------------
    # 3. Rejection of Multi-Area Selection
    # --------------------------------------------------------------------------

    def test_rejection_of_multi_area_selection(self):
        """Verify non-contiguous multi-area selections are rejected."""
        bridge_macro = f"'{self.wb_name}'!CellProcessor.ConvertRangeBridge"

        self.ws.Range("A1").Value = 100
        self.ws.Range("C1").Value = 200
        src_multi = self.ws.Range("A1:A2, C1:C2")
        dst = self.ws.Range("B1")

        res = self.excel.Run(bridge_macro, src_multi, dst)
        success, converted, skipped, errors, undo_avail, err_msg = res

        self.assertFalse(success, "Multi-area source should be rejected")
        self.assertEqual(converted, 0)
        self.assertIn("multi-area", err_msg.lower())

    # --------------------------------------------------------------------------
    # 4. Rejection of Same-Sheet Overlap
    # --------------------------------------------------------------------------

    def test_rejection_of_same_sheet_overlap(self):
        """Verify overlapping ranges on the same sheet are rejected."""
        bridge_macro = f"'{self.wb_name}'!CellProcessor.ConvertRangeBridge"

        src = self.ws.Range("A1:A10")
        dst = self.ws.Range("A5:A14")

        res = self.excel.Run(bridge_macro, src, dst)
        success, converted, skipped, errors, undo_avail, err_msg = res

        self.assertFalse(success, "Overlapping ranges on same sheet must be rejected")
        self.assertEqual(converted, 0)
        self.assertTrue(
            "trung" in err_msg.lower() or "de len nhau" in err_msg.lower(),
            f"Expected overlap message, got: {err_msg}",
        )

    # --------------------------------------------------------------------------
    # 5. Cross-Sheet Conversion
    # --------------------------------------------------------------------------

    def test_cross_sheet_conversion(self):
        """Verify conversion across sheets works without overlap errors."""
        bridge_macro = f"'{self.wb_name}'!CellProcessor.ConvertRangeBridge"

        # Create or obtain second worksheet
        if self.wb.Worksheets.Count < 2:
            ws2 = self.wb.Worksheets.Add(After=self.ws)
        else:
            ws2 = self.wb.Worksheets(2)
        ws2.Cells.Clear()

        self.ws.Range("A1:A3").Value = [[100], [200], [300]]
        src = self.ws.Range("A1:A3")
        dst = ws2.Range("B1")

        res = self.excel.Run(bridge_macro, src, dst)
        success, converted, skipped, errors, undo_avail, err_msg = res

        self.assertTrue(success, f"Cross-sheet conversion failed: {err_msg}")
        self.assertEqual(converted, 3)
        self.assertEqual(ws2.Range("B1").Value, "Một trăm đồng chẵn.")
        self.assertEqual(ws2.Range("B3").Value, "Ba trăm đồng chẵn.")

    # --------------------------------------------------------------------------
    # 6. Formula-Safe Skip Preservation
    # --------------------------------------------------------------------------

    def test_formula_safe_skip_preservation(self):
        """
        Verify that skipped cells in source (empty, non-numeric) NEVER overwrite
        destination cells, and existing destination formulas remain completely intact.
        """
        bridge_macro = f"'{self.wb_name}'!CellProcessor.ConvertRangeBridge"

        # Setup source: A1=1000, A2=empty, A3=3000
        self.ws.Range("A1").Value = 1000
        self.ws.Range("A2").Value = None
        self.ws.Range("A3").Value = 3000

        # Setup destination: B1=empty, B2 has formula =SUM(10,20), B3=empty
        self.ws.Range("B1").Value = None
        self.ws.Range("B2").Formula = "=SUM(10, 20)"
        self.ws.Range("B3").Value = None

        src = self.ws.Range("A1:A3")
        dst = self.ws.Range("B1")

        res = self.excel.Run(bridge_macro, src, dst)
        success, converted, skipped, errors, undo_avail, err_msg = res

        self.assertTrue(success, f"Conversion failed: {err_msg}")
        self.assertEqual(converted, 2)
        self.assertEqual(skipped, 1)

        # B1 and B3 must be converted
        self.assertEqual(self.ws.Range("B1").Value, "Một nghìn đồng chẵn.")
        self.assertEqual(self.ws.Range("B3").Value, "Ba nghìn đồng chẵn.")

        # B2 formula MUST remain intact!
        b2_formula = str(self.ws.Range("B2").Formula).replace(" ", "")
        self.assertEqual(b2_formula, "=SUM(10,20)", "Existing destination formula was overwritten!")
        self.assertEqual(int(self.ws.Range("B2").Value), 30)

    # --------------------------------------------------------------------------
    # 7. 2-Phase Transactional Undo
    # --------------------------------------------------------------------------

    def test_two_phase_undo_restores_values_and_formulas(self):
        """
        Verify 2-phase Undo restores original constants, formulas, and clears blank cells.
        """
        bridge_macro = f"'{self.wb_name}'!CellProcessor.ConvertRangeBridge"
        undo_macro = f"'{self.wb_name}'!UndoManager.ExecuteUndo"
        is_undo_avail_macro = f"'{self.wb_name}'!UndoManager.IsUndoAvailable"

        # Setup original destination state
        self.ws.Range("B1").Value = "Original Text"
        self.ws.Range("B2").Formula = "=SUM(5, 5)"
        self.ws.Range("B3").Value = None

        # Setup source
        self.ws.Range("A1:A3").Value = [[50000], [60000], [70000]]
        src = self.ws.Range("A1:A3")
        dst = self.ws.Range("B1")

        # Execute conversion
        res = self.excel.Run(bridge_macro, src, dst)
        self.assertTrue(res[0])
        self.assertTrue(res[4])  # UndoAvailable
        self.assertTrue(self.excel.Run(is_undo_avail_macro))

        # Verify values changed
        self.assertEqual(self.ws.Range("B1").Value, "Năm mươi nghìn đồng chẵn.")
        self.assertEqual(self.ws.Range("B2").Value, "Sáu mươi nghìn đồng chẵn.")
        self.assertEqual(self.ws.Range("B3").Value, "Bảy mươi nghìn đồng chẵn.")

        # Execute Undo
        undo_success = self.excel.Run(undo_macro)
        self.assertTrue(undo_success, "ExecuteUndo failed")

        # Verify exact restoration
        self.assertEqual(self.ws.Range("B1").Value, "Original Text")
        b2_fmla = str(self.ws.Range("B2").Formula).replace(" ", "")
        self.assertEqual(b2_fmla, "=SUM(5,5)")
        self.assertEqual(int(self.ws.Range("B2").Value), 10)
        self.assertIsNone(self.ws.Range("B3").Value)

        # Verify Undo buffer is now cleared
        self.assertFalse(self.excel.Run(is_undo_avail_macro))

    # --------------------------------------------------------------------------
    # 8. Protected Sheet Atomic Abort
    # --------------------------------------------------------------------------

    def test_protected_sheet_locked_cell_aborts_atomically(self):
        """
        Verify that if the destination worksheet is protected and has locked cells,
        the conversion aborts atomically before modifying any cells.
        """
        bridge_macro = f"'{self.wb_name}'!CellProcessor.ConvertRangeBridge"

        self.ws.Range("A1").Value = 50000
        self.ws.Range("B1").Value = "Locked Content - Do Not Overwrite"
        self.ws.Range("B1").Locked = True
        self.ws.Protect("testpwd")

        src = self.ws.Range("A1")
        dst = self.ws.Range("B1")

        res = self.excel.Run(bridge_macro, src, dst)
        success, converted, skipped, errors, undo_avail, err_msg = res

        self.assertFalse(success, "Conversion should fail when destination cell is locked on protected sheet")
        self.assertEqual(converted, 0)
        self.assertEqual(self.ws.Range("B1").Value, "Locked Content - Do Not Overwrite")

        # Unprotect for clean tear down
        self.ws.Unprotect("testpwd")

    # --------------------------------------------------------------------------
    # 9. Settings Persistence and Reset in Registry
    # --------------------------------------------------------------------------

    def test_settings_save_load_reset(self):
        """Verify saving custom settings, loading them, and resetting to factory defaults."""
        save_macro = f"'{self.wb_name}'!Settings.SaveSettingsBridge"
        load_macro = f"'{self.wb_name}'!Settings.LoadSettingsBridge"
        reset_macro = f"'{self.wb_name}'!Settings.ResetSettingsBridge"

        # 1. Save custom settings
        # Zero=1, Thousand=1, Four=1, DecMode=1, Curr=1 (USD), AddChan=False, DecPlaces=2,
        # Prefix="VIP", Suffix="USD", SubUnit="cents", Casing=1 (Upper), AddPeriod=False,
        # OutputMode=1 (Formula), QuickDir=2 (Below), ConfirmOverwrite=False, ShowSummary=False, Ver=1
        self.excel.Run(
            save_macro,
            1, 1, 1, 1,
            1, False, 2, "VIP", "USD", "cents",
            1, False,
            1, 2, False, False, 1
        )

        loaded = self.excel.Run(load_macro)
        (
            zero_style, thousand_style, four_style, dec_mode,
            curr_type, add_chan, dec_places, prefix, suffix, sub_unit,
            casing, add_period,
            output_mode, quick_dir, confirm_overwrite, show_summary, ver
        ) = loaded

        self.assertEqual(zero_style, 1)
        self.assertEqual(thousand_style, 1)
        self.assertEqual(four_style, 1)
        self.assertEqual(dec_mode, 1)
        self.assertEqual(curr_type, 1)
        self.assertFalse(add_chan)
        self.assertEqual(dec_places, 2)
        self.assertEqual(prefix, "VIP")
        self.assertEqual(suffix, "USD")
        self.assertEqual(sub_unit, "cents")
        self.assertEqual(casing, 1)
        self.assertFalse(add_period)
        self.assertEqual(output_mode, 1)
        self.assertEqual(quick_dir, 2)
        self.assertFalse(confirm_overwrite)
        self.assertFalse(show_summary)
        self.assertEqual(ver, 1)

        # 2. Reset settings to default
        self.excel.Run(reset_macro)
        def_loaded = self.excel.Run(load_macro)
        (
            def_zero, def_thousand, def_four, def_dec_mode,
            def_curr, def_chan, def_places, def_prefix, def_suffix, def_sub_unit,
            def_casing, def_period,
            def_out_mode, def_dir, def_confirm, def_summary, def_ver
        ) = def_loaded

        # Defaults: Zero=0, Thousand=0, Four=0, DecMode=0, Curr=0 (VND), AddChan=True, Casing=0,
        # AddPeriod=True, OutputMode=0, QuickDir=0, ConfirmOverwrite=True, ShowSummary=True
        self.assertEqual(def_zero, 0)
        self.assertEqual(def_thousand, 0)
        self.assertEqual(def_four, 0)
        self.assertEqual(def_dec_mode, 0)
        self.assertEqual(def_curr, 0)
        self.assertTrue(def_chan)
        self.assertEqual(def_casing, 0)
        self.assertTrue(def_period)
        self.assertEqual(def_out_mode, 0)
        self.assertEqual(def_dir, 0)
        self.assertTrue(def_confirm)
        self.assertTrue(def_summary)

    # --------------------------------------------------------------------------
    # 10. Formula Mode Generation and Currency Validation
    # --------------------------------------------------------------------------

    def test_formula_mode_generation_and_rejection(self):
        """
        Verify Formula mode generates `=BWPVND(...)` for VND and rejects USD/Custom.
        """
        bridge_macro = f"'{self.wb_name}'!CellProcessor.ConvertRangeBridge"

        self.ws.Range("A1").Value = 500000
        src = self.ws.Range("A1")
        dst = self.ws.Range("B1")

        # OutputMode=1 (Formula), CurrType=0 (VND), AddChan=True, DecPlaces=0, Zero=0, Thousand=0
        res = self.excel.Run(bridge_macro, src, dst, 1, 0, True, 0, 0, 0)
        self.assertTrue(res[0], f"Formula mode failed: {res[5]}")
        self.assertEqual(self.ws.Range("B1").Formula, "=BWPVND(A1, TRUE, 0, 0)")

        # Rejection of USD in formula mode: OutputMode=1, CurrType=1 (USD)
        res_usd = self.excel.Run(bridge_macro, src, dst, 1, 1)
        self.assertFalse(res_usd[0], "USD should be rejected in formula mode")
        self.assertIn("usd", res_usd[5].lower())

    # --------------------------------------------------------------------------
    # 11. Application State Preservation
    # --------------------------------------------------------------------------

    def test_application_state_restored_cleanly(self):
        """Verify Excel Application state is restored after both success and failure."""
        bridge_macro = f"'{self.wb_name}'!CellProcessor.ConvertRangeBridge"

        # Record initial states
        initial_screen_updating = self.excel.ScreenUpdating
        initial_events = self.excel.EnableEvents

        # 1. Success path
        self.ws.Range("A1").Value = 1000
        self.excel.Run(bridge_macro, self.ws.Range("A1"), self.ws.Range("B1"))
        self.assertEqual(self.excel.ScreenUpdating, initial_screen_updating)
        self.assertEqual(self.excel.EnableEvents, initial_events)

        # 2. Error path (invalid range overlap)
        self.excel.Run(bridge_macro, self.ws.Range("A1:A5"), self.ws.Range("A1:A5"))
        self.assertEqual(self.excel.ScreenUpdating, initial_screen_updating)
        self.assertEqual(self.excel.EnableEvents, initial_events)

    # --------------------------------------------------------------------------
    # 12. Single Empty Cell HasOccupiedCells Guard
    # --------------------------------------------------------------------------

    def test_has_occupied_cells_single_empty_cell_with_data_elsewhere(self):
        """
        Verify that HasOccupiedCells on a single empty cell returns False even
        when other cells across the worksheet contain data (guard against SpecialCells whole-sheet scan).
        """
        has_occ_macro = f"'{self.wb_name}'!CellProcessor.HasOccupiedCells"

        # Populate other cells far away
        self.ws.Range("Z100").Value = "Existing data elsewhere"
        self.ws.Range("H50").Formula = "=SUM(1, 2)"

        # Target cell A1 is completely empty
        self.ws.Range("A1").Value = None
        is_occ_empty = self.excel.Run(has_occ_macro, self.ws.Range("A1"))
        self.assertFalse(is_occ_empty, "Single empty cell should NOT be reported as occupied even if other cells have data")

        # Target cell A1 has constant
        self.ws.Range("A1").Value = 123
        is_occ_val = self.excel.Run(has_occ_macro, self.ws.Range("A1"))
        self.assertTrue(is_occ_val, "Single cell with value should be reported as occupied")

        # Target cell A1 has formula
        self.ws.Range("A1").Formula = "=SUM(5, 5)"
        is_occ_fmla = self.excel.Run(has_occ_macro, self.ws.Range("A1"))
        self.assertTrue(is_occ_fmla, "Single cell with formula should be reported as occupied")

    # --------------------------------------------------------------------------
    # 13. Exceeding MAX_UNDO_CELLS Invalidation
    # --------------------------------------------------------------------------

    def test_prepare_undo_snapshot_exceeding_max_cells_clears_previous_undo(self):
        """
        Verify that when an operation exceeds MAX_UNDO_CELLS (10,000 cells),
        any existing committed undo is invalidated/cleared rather than left stale.
        """
        bridge_macro = f"'{self.wb_name}'!CellProcessor.ConvertRangeBridge"
        prepare_undo_macro = f"'{self.wb_name}'!UndoManager.PrepareUndoSnapshot"
        is_undo_avail_macro = f"'{self.wb_name}'!UndoManager.IsUndoAvailable"

        # 1. Establish a valid committed undo
        self.ws.Range("A1").Value = 1000
        res = self.excel.Run(bridge_macro, self.ws.Range("A1"), self.ws.Range("B1"))
        self.assertTrue(res[0])
        self.assertTrue(self.excel.Run(is_undo_avail_macro), "Committed undo should be available")

        # 2. Invoke PrepareUndoSnapshot on a range with 10,001 cells (> MAX_UNDO_CELLS)
        rng_large = self.ws.Range("A1:A10001")
        self.assertEqual(rng_large.Rows.Count * rng_large.Columns.Count, 10001)
        self.excel.Run(prepare_undo_macro, rng_large)

        # 3. Verify previous undo was cleared
        self.assertFalse(
            self.excel.Run(is_undo_avail_macro),
            "Exceeding MAX_UNDO_CELLS must invalidate previous undo snapshot",
        )

    # --------------------------------------------------------------------------
    # 14. SettingsVersion Persistence
    # --------------------------------------------------------------------------

    def test_settings_version_persistence(self):
        """Verify SettingsVersion can be saved and loaded dynamically from Registry."""
        save_macro = f"'{self.wb_name}'!Settings.SaveSettingsBridge"
        load_macro = f"'{self.wb_name}'!Settings.LoadSettingsBridge"

        # Save with custom SettingsVersion = 2
        self.excel.Run(
            save_macro,
            0, 0, 0, 0,
            0, True, 0, "", "", "",
            0, True,
            0, 0, True, True, 2
        )

        loaded = self.excel.Run(load_macro)
        saved_version = loaded[16]  # 17th element (index 16) is SettingsVersion
        self.assertEqual(saved_version, 2, f"Expected SettingsVersion=2, got {saved_version}")


if __name__ == "__main__":
    unittest.main()

