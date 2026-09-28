"""
BWPConvertTTNVN Build Orchestration & Compilation Pipeline
Provides version injection, manifest validation, safe Excel COM process isolation with PID tracking,
direct OpenXML Ribbon injection, reopen validation, and automated test execution.
"""

import argparse
import gc
import os
import re
import subprocess
import sys
import time
import winreg
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple, Union

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

try:
    import pythoncom
    import win32api
    import win32com.client
    import win32con
    import win32process
except ImportError:
    pass  # Allow import on non-Windows platforms for doc generation / static checks

# Project Constants
APP_NAME = "BWPConvertTTNVN"
APP_AUTHOR = "IT Leon"
APP_SUBJECT = "Vietnamese Number to Words Excel Add-in"
APP_COPYRIGHT = "Copyright © 2026 - IT Leon"
APP_COMMENTS = "Copyright (c) 2026 - IT Leon"

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
DEFAULT_VERSION_FILE = PROJECT_ROOT / "VERSION"
DEFAULT_BUILD_INFO_FILE = PROJECT_ROOT / "src" / "excel" / "BuildInfo.bas"
DEFAULT_XLAM_PATH = PROJECT_ROOT / "dist" / "BWPConvertTTNVN.xlam"
DEFAULT_RIBBON_XML = PROJECT_ROOT / "ribbon" / "customUI14.xml"
DEFAULT_TEST_REPORT = PROJECT_ROOT / "dist" / "test-report.txt"

# Strict 15-Component Manifest Order
MANIFEST_COMPONENTS: List[Tuple[str, str, bool]] = [
    ("src/core", "CoreTypes.bas", False),
    ("src/core", "UnicodeText.bas", False),
    ("src/core", "VietnameseNumber.bas", False),
    ("src/core", "Currency.bas", False),
    ("src/core", "TextFormatter.bas", False),
    ("src/core", "CoreCoordinator.bas", False),
    ("src/excel", "BuildInfo.bas", False),
    ("src/excel", "Settings.bas", False),
    ("src/excel", "UndoManager.bas", False),
    ("src/excel", "CellProcessor.bas", False),
    ("src/excel", "UDF.bas", False),
    ("src/excel", "RibbonCallbacks.bas", False),
    ("src/forms", "frmConvert.frm", True),
    ("src/forms", "frmSettings.frm", True),
    ("src/forms", "frmAbout.frm", True),
]

# Expected VBComponent Names inside compiled VBProject
EXPECTED_VB_COMPONENTS: Set[str] = {
    "CoreTypes",
    "UnicodeText",
    "VietnameseNumber",
    "VnCurrency",
    "TextFormatter",
    "CoreCoordinator",
    "BuildInfo",
    "Settings",
    "UndoManager",
    "CellProcessor",
    "UDF",
    "RibbonCallbacks",
    "frmConvert",
    "frmSettings",
    "frmAbout",
}


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
        except Exception as e:
            print(f"[WARN] Failed to resolve Excel HWND/PID: {e}", file=sys.stderr)
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
                return exit_code == win32con.STILL_ACTIVE
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
    print("=" * 70)
    print("  BWPConvertTTNVN - Build Preflight Diagnostics")
    print("=" * 70)
    for r in summary["results"]:
        status = "[ OK ]" if r["ok"] else "[FAIL]"
        print(f" {status} {r['check']}: {r['detail']}")
    print("-" * 70)
    if summary["all_passed"]:
        print(" [SUCCESS] All preflight checks passed. Environment is ready to build.")
    else:
        print(" [WARNING] One or more preflight checks failed.")
    print("=" * 70)


# ---------------------------------------------------------------------------
# Multi-Stage Build Pipeline
# ---------------------------------------------------------------------------

def stage_0_version_and_manifest(project_root: Path) -> str:
    """
    Stage 0: Version & Manifest Check:
    - Reads VERSION and synchronizes src/excel/BuildInfo.bas.
    - Validates presence and 7-bit ASCII safety of all 15 manifest components.
    - Verifies companion .frx binary files exist for all UserForms.
    - Verifies customUI14.xml callbacks match public procedures in RibbonCallbacks.bas.
    """
    print("\n--- [Stage 0: Version & Manifest Check] ---")

    # 1. Version sync
    version_str = read_version(project_root / "VERSION")
    build_info_path = generate_build_info(version_str, project_root / "src" / "excel" / "BuildInfo.bas")
    print(f"  [Version] Synced version {version_str} -> {build_info_path.relative_to(project_root)}")

    # 2. Manifest check
    print(f"  [Manifest] Validating {len(MANIFEST_COMPONENTS)} components in strict manifest order...")
    for sub_dir, filename, has_frx in MANIFEST_COMPONENTS:
        src_file = project_root / sub_dir / filename
        if not src_file.is_file():
            raise FileNotFoundError(f"Manifest component missing: {src_file}")

        # Check pure 7-bit ASCII safety for code files (excluding BuildInfo.bas which contains copyright symbol)
        if filename != "BuildInfo.bas":
            content = src_file.read_bytes()
            non_ascii = [b for b in content if b > 127]
            if non_ascii:
                raise ValueError(
                    f"Non-ASCII bytes detected in {src_file.name} ({len(non_ascii)} occurrences). "
                    f"All VBA code must be pure 7-bit ASCII safe."
                )

        if has_frx:
            frx_file = src_file.with_suffix(".frx")
            if not frx_file.is_file():
                raise FileNotFoundError(f"Companion binary form file missing: {frx_file}")
            if frx_file.stat().st_size == 0:
                raise ValueError(f"Companion binary form file is empty: {frx_file}")

        print(f"    [OK] {sub_dir}/{filename}")

    # 3. Ribbon Callback Contract Check
    ribbon_xml_path = project_root / "ribbon" / "customUI14.xml"
    ribbon_bas_path = project_root / "src" / "excel" / "RibbonCallbacks.bas"
    if not ribbon_xml_path.is_file():
        raise FileNotFoundError(f"Ribbon XML missing: {ribbon_xml_path}")
    if not ribbon_bas_path.is_file():
        raise FileNotFoundError(f"RibbonCallbacks.bas missing: {ribbon_bas_path}")

    ribbon_bytes = ribbon_xml_path.read_bytes()
    if ribbon_bytes.startswith(b"\xef\xbb\xbf"):
        ribbon_bytes = ribbon_bytes[3:]
    ribbon_tree = ET.fromstring(ribbon_bytes)

    callbacks_in_xml = set()
    for elem in ribbon_tree.iter():
        for attr in ["onLoad", "onAction", "getEnabled"]:
            cb = elem.attrib.get(attr)
            if cb:
                callbacks_in_xml.add(cb)

    bas_text = ribbon_bas_path.read_text(encoding="ascii")
    for cb in callbacks_in_xml:
        pattern = rf"\bPublic\s+(?:Sub|Function|Property\s+Get)\s+{re.escape(cb)}\b"
        if not re.search(pattern, bas_text, re.IGNORECASE):
            raise ValueError(f"Ribbon callback '{cb}' defined in XML but missing in {ribbon_bas_path.name}")

    print(f"  [Ribbon] Validated {len(callbacks_in_xml)} Ribbon callbacks contract against RibbonCallbacks.bas.")
    print("  [Stage 0: PASSED]")
    return version_str


def stage_1_preflight_check() -> Dict[str, Any]:
    """
    Stage 1: Environment Preflight:
    - Verifies Windows OS, Excel registry, AccessVBOM setting, and live COM capability.
    """
    print("\n--- [Stage 1: Environment Preflight] ---")
    summary = run_preflight()
    print_preflight_report(summary)
    if not summary["all_passed"]:
        raise RuntimeError("Environment preflight checks failed! AccessVBOM=1 and Excel COM are required to build.")
    print("  [Stage 1: PASSED]")
    return summary


def stage_2_compile_xlam(project_root: Path, target_xlam: Path) -> Path:
    """
    Stage 2: Isolated COM Build:
    - Launches isolated Excel COM instance (DispatchEx) with PID tracking.
    - Creates a blank workbook.
    - Injects workbook document metadata (Title, Author, Subject, Comments).
    - Imports all 15 components in strict manifest order.
    - Saves workbook as .xlam (FileFormat = 55, xlOpenXMLAddIn).
    - Cleanly closes workbook, quits Excel, guarantees process cleanup.
    """
    print("\n--- [Stage 2: Isolated COM Compilation] ---")
    target_xlam.parent.mkdir(parents=True, exist_ok=True)
    if target_xlam.is_file():
        try:
            target_xlam.unlink()
        except Exception as e:
            raise IOError(f"Cannot overwrite existing add-in at {target_xlam}: {e}")

    with ExcelSession(visible=False, display_alerts=False) as session:
        print(f"  [Excel COM] Started isolated instance (PID: {session.pid})")
        wb = session.excel.Workbooks.Add()

        # Set workbook metadata
        try:
            wb.BuiltinDocumentProperties("Title").Value = APP_NAME
            wb.BuiltinDocumentProperties("Author").Value = APP_AUTHOR
            wb.BuiltinDocumentProperties("Subject").Value = APP_SUBJECT
            wb.BuiltinDocumentProperties("Comments").Value = APP_COMMENTS
            print("  [Metadata] Injected Title, Author, Subject, and Comments.")
        except Exception as e:
            print(f"  [WARN] Failed to set document properties: {e}")

        # Import 15 components in strict order
        vb_proj = wb.VBProject
        print(f"  [VBA Import] Importing {len(MANIFEST_COMPONENTS)} components in strict manifest order...")
        for sub_dir, filename, _ in MANIFEST_COMPONENTS:
            file_path = project_root / sub_dir / filename
            vb_proj.VBComponents.Import(str(file_path))
            print(f"    -> Imported: {filename}")

        # Save as xlOpenXMLAddIn (FileFormat 55)
        print(f"  [SaveAs] Saving add-in to: {target_xlam} (FileFormat=55, xlOpenXMLAddIn)")
        wb.SaveAs(Filename=str(target_xlam), FileFormat=55)
        wb.Close(SaveChanges=False)

    print("  [Process] Excel COM instance exited and PID cleaned up.")
    if not target_xlam.is_file():
        raise FileNotFoundError(f"Failed to generate add-in: {target_xlam}")
    print(f"  [Stage 2: PASSED] Add-in created ({target_xlam.stat().st_size:,} bytes)")
    return target_xlam


def stage_3_inject_ribbon(target_xlam: Path, ribbon_xml: Path) -> Path:
    """
    Stage 3: Direct OpenXML Ribbon Injection:
    - Calls package_ribbon.inject_ribbon directly on target_xlam.
    """
    print("\n--- [Stage 3: Direct OpenXML Ribbon Injection] ---")
    from scripts.package_ribbon import inject_ribbon

    out = inject_ribbon(target_xlam, ribbon_xml)
    print(f"  [Ribbon] Injected customUI14.xml into: {out}")
    print("  [Stage 3: PASSED]")
    return out


def stage_4_reopen_validation(target_xlam: Path):
    """
    Stage 4: Reopen Validation:
    - Fresh isolated Excel COM instance opens target_xlam.
    - Verifies OpenXML structural validity and that all 15 VBA components load without error.
    - Executes a smoke test on the core coordinator macro.
    """
    print("\n--- [Stage 4: Reopen Validation] ---")
    with ExcelSession(visible=False, display_alerts=False) as session:
        print(f"  [Excel COM] Opening packaged add-in in fresh instance (PID: {session.pid})...")
        wb = session.excel.Workbooks.Open(str(target_xlam))
        if wb is None:
            raise RuntimeError(f"Excel failed to open packaged add-in: {target_xlam}")

        # Check all 15 VBA components exist
        actual_comps = {c.Name for c in wb.VBProject.VBComponents}
        missing = EXPECTED_VB_COMPONENTS - actual_comps
        if missing:
            raise RuntimeError(f"Reopened add-in is missing expected VBA components: {missing}")
        print(f"  [VBA Check] Verified all {len(EXPECTED_VB_COMPONENTS)} components present in VBProject.")

        # Smoke test execution
        smoke_val = 125430000
        smoke_macro = f"'{wb.Name}'!ConvertNumberDefault"
        res = session.excel.Run(smoke_macro, smoke_val)
        expected = "Một trăm hai mươi lăm triệu bốn trăm ba mươi nghìn đồng chẵn."
        if res != expected:
            raise ValueError(f"Smoke test returned unexpected result: '{res}' != '{expected}'")
        print(f"  [Smoke Test] Verified '{smoke_macro}'({smoke_val:,}) -> '{res}'")

        wb.Close(SaveChanges=False)

    print("  [Stage 4: PASSED]")


def stage_5_run_tests(target_xlam: Path, report_path: Path):
    """
    Stage 5: Test Execution:
    - Invokes tests/run_tests.py via subprocess to run the 3-tier test suite.
    - Generates dist/test-report.txt.
    """
    print("\n--- [Stage 5: Automated 3-Tier Test Execution] ---")
    test_runner_script = PROJECT_ROOT / "tests" / "run_tests.py"
    if not test_runner_script.is_file():
        raise FileNotFoundError(f"Test runner script missing: {test_runner_script}")

    report_path.parent.mkdir(parents=True, exist_ok=True)

    cmd = [
        sys.executable,
        str(test_runner_script),
        "--xlam",
        str(target_xlam),
        "--report",
        str(report_path),
    ]

    print(f"  [Runner] Executing: {' '.join(cmd)}")
    proc = subprocess.run(cmd, check=False)
    if proc.returncode != 0:
        raise RuntimeError(
            f"3-Tier test suite failed with exit code {proc.returncode}. "
            f"Check report at: {report_path}"
        )
    print("  [Stage 5: PASSED]")


# ---------------------------------------------------------------------------
# Main Build Orchestrator
# ---------------------------------------------------------------------------

def run_build(
    skip_tests: bool = False,
    output_xlam: Optional[Union[str, Path]] = None,
    ribbon_xml: Optional[Union[str, Path]] = None,
    report_path: Optional[Union[str, Path]] = None,
) -> Path:
    """
    Executes the full automated build pipeline stages 0 through 5.
    """
    start_time = time.time()
    target_xlam = Path(output_xlam).resolve() if output_xlam else DEFAULT_XLAM_PATH
    target_ribbon = Path(ribbon_xml).resolve() if ribbon_xml else DEFAULT_RIBBON_XML
    target_report = Path(report_path).resolve() if report_path else DEFAULT_TEST_REPORT

    print("=" * 70)
    print(f"  {APP_NAME} - Automated Build Pipeline")
    print(f"  Target: {target_xlam}")
    print("=" * 70)

    # Stage 0: Version & Manifest Check
    stage_0_version_and_manifest(PROJECT_ROOT)

    # Stage 1: Environment Preflight
    stage_1_preflight_check()

    # Stage 2: Isolated COM Build
    stage_2_compile_xlam(PROJECT_ROOT, target_xlam)

    # Stage 3: Direct OpenXML Ribbon Injection
    stage_3_inject_ribbon(target_xlam, target_ribbon)

    # Stage 4: Reopen Validation
    stage_4_reopen_validation(target_xlam)

    # Stage 5: Test Execution
    if not skip_tests:
        stage_5_run_tests(target_xlam, target_report)
    else:
        print("\n--- [Stage 5: Skipped by Flag] ---")

    elapsed = time.time() - start_time
    print("\n" + "=" * 70)
    print(f"  [BUILD SUCCESS] {APP_NAME} compiled successfully in {elapsed:.2f}s!")
    print(f"  Add-in artifact: {target_xlam}")
    if not skip_tests and target_report.is_file():
        print(f"  Test report:     {target_report}")
    print("=" * 70)
    return target_xlam


def main():
    parser = argparse.ArgumentParser(description="BWPConvertTTNVN Build Orchestration & Compilation Tool")
    parser.add_argument("--preflight", action="store_true", help="Run preflight environment diagnostics only")
    parser.add_argument("--generate-build-info", action="store_true", help="Regenerate src/excel/BuildInfo.bas from VERSION only")
    parser.add_argument("--skip-tests", action="store_true", help="Skip running the 3-tier test suite in Stage 5")
    parser.add_argument("--output", default=str(DEFAULT_XLAM_PATH), help=f"Destination path for .xlam (default: {DEFAULT_XLAM_PATH})")
    parser.add_argument("--ribbon-xml", default=str(DEFAULT_RIBBON_XML), help=f"Path to customUI XML (default: {DEFAULT_RIBBON_XML})")
    parser.add_argument("--test-report", default=str(DEFAULT_TEST_REPORT), help=f"Path to test report output (default: {DEFAULT_TEST_REPORT})")

    args = parser.parse_args()

    if args.generate_build_info and not (args.preflight or args.skip_tests):
        version = read_version()
        out = generate_build_info(version)
        print(f"[BuildInfo] Generated {out} with version {version}")
        return

    if args.preflight:
        summary = run_preflight()
        print_preflight_report(summary)
        if not summary["all_passed"]:
            sys.exit(1)
        return

    # Default: Run complete multi-stage build pipeline
    try:
        run_build(
            skip_tests=args.skip_tests,
            output_xlam=args.output,
            ribbon_xml=args.ribbon_xml,
            report_path=args.test_report,
        )
    except Exception as e:
        print(f"\n[BUILD FAILED] {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
