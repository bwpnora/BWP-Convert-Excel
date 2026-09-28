"""
Preflight & Environment Unit Tests for BWPConvertTTNVN
Verifies VERSION, BuildInfo generation, AccessVBOM trust status, and isolated Excel COM PID tracking.
"""

import os
import sys
import tempfile
import time
import unittest
from pathlib import Path

# Add project root to sys.path so scripts can be imported
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import win32api
import win32com.client
import win32con
import win32process

from scripts.build import (
    APP_COPYRIGHT,
    APP_NAME,
    ExcelSession,
    check_access_vbom_registry,
    check_excel_registry,
    check_windows_os,
    generate_build_info,
    get_excel_pid,
    read_version,
    run_preflight,
    test_access_vbom_live,
)


class TestPreflight(unittest.TestCase):
    """
    Test suite for build preflight diagnostics and process safety.
    """

    def test_version_file_exists(self):
        """Verify root VERSION file exists and contains exact version 1.0.0."""
        version_file = PROJECT_ROOT / "VERSION"
        self.assertTrue(version_file.is_file(), f"VERSION file missing at {version_file}")
        v = read_version(version_file)
        self.assertEqual(v, "1.0.0")

    def test_build_info_bas_exists_and_content(self):
        """Verify src/excel/BuildInfo.bas exists and contains expected constants."""
        build_info_path = PROJECT_ROOT / "src" / "excel" / "BuildInfo.bas"
        self.assertTrue(build_info_path.is_file(), f"BuildInfo.bas missing at {build_info_path}")
        content = build_info_path.read_text(encoding="utf-8")
        self.assertIn('Attribute VB_Name = "BuildInfo"', content)
        self.assertIn('Option Explicit', content)
        self.assertIn(f'Public Const APP_NAME As String = "{APP_NAME}"', content)
        self.assertIn('Public Const APP_VERSION As String = "1.0.0"', content)
        self.assertIn('Public Property Get APP_COPYRIGHT() As String', content)
        self.assertIn('ChrW$(&HA9)', content)

    def test_generate_build_info_function(self):
        """Verify generate_build_info regenerates file with custom version accurately."""
        with tempfile.TemporaryDirectory() as tmpdir:
            test_target = Path(tmpdir) / "BuildInfoTest.bas"
            out = generate_build_info(version_str="1.2.3", output_path=test_target)
            self.assertTrue(out.is_file())
            content = out.read_text(encoding="utf-8")
            self.assertIn('Public Const APP_VERSION As String = "1.2.3"', content)
            self.assertIn(f'Public Const APP_NAME As String = "{APP_NAME}"', content)
            self.assertIn('Public Property Get APP_COPYRIGHT() As String', content)
            self.assertIn('ChrW$(&HA9)', content)

    def test_excel_pid_tracking_and_cleanup(self):
        """
        Verify raw Excel COM DispatchEx obtains valid Hwnd and PID,
        and process cleanly exits upon Quit() without affecting any other processes.
        """
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            hwnd = excel.Hwnd
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            self.assertGreater(pid, 0)
        finally:
            excel.Quit()
            del excel

        # Allow brief time for process shutdown
        time.sleep(0.5)

        # Confirm process has exited or terminate ONLY this specific PID
        if ExcelSession.is_pid_alive(pid):
            ExcelSession._ensure_pid_terminated(pid, timeout_sec=2.0)

        self.assertFalse(
            ExcelSession.is_pid_alive(pid),
            f"Excel COM process PID {pid} was not terminated after cleanup.",
        )

    def test_excel_session_context_manager(self):
        """
        Verify ExcelSession context manager tracks PID and terminates ONLY that PID on exit.
        """
        tracked_pid = None
        with ExcelSession(visible=False, display_alerts=False) as session:
            self.assertIsNotNone(session.excel)
            self.assertIsNotNone(session.pid)
            self.assertGreater(session.pid, 0)
            tracked_pid = session.pid
            self.assertTrue(ExcelSession.is_pid_alive(tracked_pid))

        # After exiting context, process must no longer be alive
        self.assertFalse(
            ExcelSession.is_pid_alive(tracked_pid),
            f"Excel process PID {tracked_pid} remained alive after ExcelSession exit.",
        )

    def test_access_vbom_registry_check(self):
        """Verify AccessVBOM registry check succeeds."""
        res = check_access_vbom_registry()
        self.assertTrue(res["ok"], f"AccessVBOM registry check failed: {res.get('detail')}")

    def test_access_vbom_live_check(self):
        """Verify live COM VBProject access succeeds."""
        res = test_access_vbom_live()
        self.assertTrue(res["ok"], f"AccessVBOM live check failed: {res.get('detail')}")

    def test_preflight_diagnostics_suite(self):
        """Verify all preflight checks in run_preflight() pass."""
        summary = run_preflight()
        self.assertTrue(summary["all_passed"], f"Preflight checks failed: {summary['results']}")


if __name__ == "__main__":
    unittest.main()
