"""
BWPConvertTTNVN Build Orchestration & Preflight System
Provides version injection, registry validation, safe Excel COM process isolation with PID tracking.
"""

import gc
import os
import subprocess
import sys
import time
import winreg
from pathlib import Path
from typing import Any, Dict, Optional, Tuple, Union

try:
    import pythoncom
    import win32api
    import win32com.client
    import win32con
    import win32process
except ImportError:
    pass  # Allow import on non-Windows platforms for testing / doc generation

# Project Constants
APP_NAME = "BWPConvertTTNVN"
APP_COPYRIGHT = "Copyright © 2026 - IT Leon"

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_VERSION_FILE = PROJECT_ROOT / "VERSION"
DEFAULT_BUILD_INFO_FILE = PROJECT_ROOT / "src" / "excel" / "BuildInfo.bas"


# ---------------------------------------------------------------------------
# Version Management & BuildInfo Generation
# ---------------------------------------------------------------------------

def read_version(version_path: Optional[Union[str, Path]] = None) -> str:
    """
    Reads the single source of truth version from the VERSION file.
    """
    path = Path(version_path) if version_path else DEFAULT_VERSION_FILE
    if not path.is_file():
        raise FileNotFoundError(f"VERSION file not found at: {path}")
    version_str = path.read_text(encoding="utf-8").strip()
    if not version_str:
        raise ValueError(f"VERSION file at {path} is empty")
    return version_str


def generate_build_info(
    version_str: Optional[str] = None,
    output_path: Optional[Union[str, Path]] = None,
) -> Path:
    """
    Generates src/excel/BuildInfo.bas using the version string.
    Ensures that the VBA layer is always synchronized with the VERSION file.
    """
    if version_str is None:
        version_str = read_version()

    target = Path(output_path) if output_path else DEFAULT_BUILD_INFO_FILE
    target.parent.mkdir(parents=True, exist_ok=True)

    content = (
        'Attribute VB_Name = "BuildInfo"\n'
        "Option Explicit\n"
        "\n"
        f'Public Const APP_NAME As String = "{APP_NAME}"\n'
        f'Public Const APP_VERSION As String = "{version_str}"\n'
        f'Public Const APP_COPYRIGHT As String = "{APP_COPYRIGHT}"\n'
    )

    target.write_text(content, encoding="utf-8")
    return target


# ---------------------------------------------------------------------------
# Excel COM Automation & Strict PID Tracking
# ---------------------------------------------------------------------------

def get_excel_pid(excel_app: Any) -> int:
    """
    Extracts the Windows Process ID (PID) from an Excel.Application COM instance.
    Uses excel_app.Hwnd and win32process.GetWindowThreadProcessId.
    """
    hwnd = excel_app.Hwnd
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    return pid


class ExcelSession:
    """
    Context manager providing safe, isolated Excel COM automation with strict PID tracking.
    Guarantees:
    1. Excel is launched isolated (DispatchEx).
    2. The exact process ID (PID) of THIS instance is tracked via its window handle.
    3. On exit, workbooks are closed, Excel is quit, COM references are released.
    4. If the process remains alive after a grace period, ONLY this specific PID is terminated.
    5. Never touches or terminates any other Excel process on the system.
    """

    def __init__(self, visible: bool = False, display_alerts: bool = False):
        self.visible = visible
        self.display_alerts = display_alerts
        self.excel = None
        self.pid: Optional[int] = None
        self.hwnd: Optional[int] = None

    def __enter__(self) -> "ExcelSession":
        if sys.platform != "win32":
            raise RuntimeError("Excel COM automation requires Windows OS.")

        self.excel = win32com.client.DispatchEx("Excel.Application")
        self.excel.Visible = self.visible
        self.excel.DisplayAlerts = self.display_alerts

        try:
            self.hwnd = self.excel.Hwnd
            self.pid = get_excel_pid(self.excel)
        except Exception:
            self.pid = None

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def close(self):
        """
        Cleanly closes workbooks, shuts down Excel, and ensures the tracked PID is terminated.
        """
        tracked_pid = self.pid

        if self.excel is not None:
            try:
                # Close any open workbooks without saving
                count = self.excel.Workbooks.Count
                for i in range(count, 0, -1):
                    try:
                        self.excel.Workbooks(i).Close(SaveChanges=False)
                    except Exception:
                        pass
            except Exception:
                pass

            try:
                self.excel.Quit()
            except Exception:
                pass

            # Explicitly release COM pointer and force COM garbage collection
            self.excel = None
            try:
                pythoncom.CoFreeUnusedLibraries()
            except Exception:
                pass
            gc.collect()

        # Check and ensure ONLY this specific PID is terminated
        if tracked_pid and tracked_pid > 0:
            self._ensure_pid_terminated(tracked_pid)
            self.pid = None

    @staticmethod
    def is_pid_alive(pid: int) -> bool:
        """
        Checks whether a specific process ID is currently running.
        """
        if not pid or pid <= 0:
            return False
        try:
            handle = win32api.OpenProcess(win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
            if handle:
                exit_code = win32process.GetExitCodeProcess(handle)
                win32api.CloseHandle(handle)
                return exit_code == 259  # STILL_ACTIVE
            return False
        except Exception:
            return False

    @classmethod
    def _ensure_pid_terminated(cls, pid: int, timeout_sec: float = 2.0):
        """
        Waits up to timeout_sec for the process to exit. If still alive, terminates ONLY this PID.
        """
        start_time = time.time()
        while time.time() - start_time < timeout_sec:
            if not cls.is_pid_alive(pid):
                return
            time.sleep(0.1)

        # Force terminate ONLY this specific PID if it failed to exit cleanly
        if cls.is_pid_alive(pid):
            terminated = False
            try:
                handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
                if handle:
                    win32process.TerminateProcess(handle, 0)
                    win32api.CloseHandle(handle)
                    terminated = True
            except Exception:
                pass

            if not terminated:
                try:
                    subprocess.run(
                        ["taskkill", "/F", "/PID", str(pid)],
                        capture_output=True,
                        check=False,
                    )
                except Exception:
                    pass


# ---------------------------------------------------------------------------
# Preflight Diagnostics
# ---------------------------------------------------------------------------

def check_windows_os() -> Dict[str, Any]:
    """
    Verifies the build is running on Windows.
    """
    is_windows = sys.platform == "win32" or os.name == "nt"
    return {
        "check": "Windows OS",
        "ok": is_windows,
        "detail": f"Platform: {sys.platform}, os.name: {os.name}",
    }


def check_excel_registry() -> Dict[str, Any]:
    """
    Verifies that Microsoft Excel is registered in Windows COM registry.
    """
    if sys.platform != "win32":
        return {"check": "Excel Registry", "ok": False, "detail": "Non-Windows platform"}

    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application") as key:
            curver = ""
            try:
                with winreg.OpenKey(key, "CurVer") as curkey:
                    curver, _ = winreg.QueryValueEx(curkey, "")
            except Exception:
                pass
            return {
                "check": "Excel COM Registry",
                "ok": True,
                "detail": f"Excel.Application found ({curver or 'Default Version'})",
            }
    except Exception as e:
        return {
            "check": "Excel COM Registry",
            "ok": False,
            "detail": f"Excel.Application not registered in COM: {e}",
        }


def check_access_vbom_registry() -> Dict[str, Any]:
    """
    Checks the AccessVBOM setting in HKCU registry across Office versions.
    AccessVBOM=1 allows programmatic access to the VBA project object model.
    """
    if sys.platform != "win32":
        return {"check": "AccessVBOM Registry", "ok": False, "detail": "Non-Windows platform"}

    known_versions = ["16.0", "15.0", "14.0", "12.0"]
    found_versions = []

    for ver in known_versions:
        reg_path = rf"Software\Microsoft\Office\{ver}\Excel\Security"
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path) as key:
                try:
                    val, _ = winreg.QueryValueEx(key, "AccessVBOM")
                    if val == 1:
                        return {
                            "check": "AccessVBOM Registry",
                            "ok": True,
                            "detail": f"AccessVBOM is enabled (1) in Office {ver} ({reg_path})",
                            "version": ver,
                            "value": val,
                        }
                    else:
                        found_versions.append(f"Office {ver}: AccessVBOM={val}")
                except FileNotFoundError:
                    found_versions.append(f"Office {ver}: AccessVBOM value not set")
        except FileNotFoundError:
            continue

    return {
        "check": "AccessVBOM Registry",
        "ok": False,
        "detail": (
            "AccessVBOM not enabled in registry. "
            + (f"Checked: {', '.join(found_versions)}" if found_versions else "No Excel Security keys found")
        ),
    }


def test_access_vbom_live() -> Dict[str, Any]:
    """
    Performs a live COM capability test: launches an isolated Excel instance, creates a temporary
    workbook, and tests reading wb.VBProject.Name to verify trust permissions.
    """
    if sys.platform != "win32":
        return {"check": "AccessVBOM Live Capability", "ok": False, "detail": "Non-Windows platform"}

    try:
        with ExcelSession(visible=False, display_alerts=False) as session:
            wb = session.excel.Workbooks.Add()
            try:
                project_name = wb.VBProject.Name
                return {
                    "check": "AccessVBOM Live Capability",
                    "ok": True,
                    "detail": f"Live COM VBProject access succeeded (Default Name: '{project_name}', Excel PID: {session.pid})",
                }
            except Exception as e:
                return {
                    "check": "AccessVBOM Live Capability",
                    "ok": False,
                    "detail": f"Live COM VBProject access rejected: {e}",
                }
            finally:
                wb.Close(SaveChanges=False)
    except Exception as e:
        return {
            "check": "AccessVBOM Live Capability",
            "ok": False,
            "detail": f"Failed to initialize Excel COM instance: {e}",
        }


def run_preflight() -> Dict[str, Any]:
    """
    Executes all preflight diagnostics and returns a comprehensive summary.
    """
    results = [
        check_windows_os(),
        check_excel_registry(),
        check_access_vbom_registry(),
        test_access_vbom_live(),
    ]

    all_passed = all(r["ok"] for r in results)

    return {
        "all_passed": all_passed,
        "results": results,
    }


def print_preflight_report(summary: Dict[str, Any]):
    """
    Prints a formatted console report of the preflight diagnostics.
    """
    print("=" * 65)
    print("  BWPConvertTTNVN - Build Preflight Diagnostics")
    print("=" * 65)
    for r in summary["results"]:
        status = "[ OK ]" if r["ok"] else "[FAIL]"
        print(f" {status} {r['check']}: {r['detail']}")
    print("-" * 65)
    if summary["all_passed"]:
        print(" [SUCCESS] All preflight checks passed. Environment is ready to build.")
    else:
        print(" [WARNING] One or more preflight checks failed.")
    print("=" * 65)


# ---------------------------------------------------------------------------
# CLI Entry Point
# ---------------------------------------------------------------------------

def main():
    import argparse

    parser = argparse.ArgumentParser(description="BWPConvertTTNVN Build & Preflight Tool")
    parser.add_argument("--preflight", action="store_true", help="Run preflight environment diagnostics")
    parser.add_argument("--generate-build-info", action="store_true", help="Regenerate src/excel/BuildInfo.bas from VERSION")

    args = parser.parse_args()

    # If no flags passed, default to preflight and build info generation
    if not (args.preflight or args.generate_build_info):
        args.preflight = True
        args.generate_build_info = True

    if args.generate_build_info:
        version = read_version()
        out = generate_build_info(version)
        print(f"[BuildInfo] Generated {out} with version {version}")

    if args.preflight:
        summary = run_preflight()
        print_preflight_report(summary)
        if not summary["all_passed"]:
            sys.exit(1)


if __name__ == "__main__":
    main()
