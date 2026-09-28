"""
Unit and Integration Tests for BWPConvertTTNVN Release Packaging & Installer
Verifies:
1. Mandatory copyright, versioning, and documentation integrity (LICENSE, CHANGELOG.md, README.md).
2. Installer contract and syntax (installer/Install.vbs).
3. Stage 6 release packaging pipeline (.zip archive and .sha256 generation).
4. Cryptographic integrity: SHA-256 verification against actual .xlam artifact.
5. ZIP bundle content completeness and structural validity.
"""

import hashlib
import os
import re
import subprocess
import sys
import unittest
import zipfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from scripts.build import (
    APP_NAME,
    DEFAULT_XLAM_PATH,
    read_version,
    stage_6_release_packaging,
)


class TestDocumentationAndMetadata(unittest.TestCase):
    """
    Validates license, documentation, and versioning single source of truth.
    """

    def setUp(self):
        self.version_str = read_version(PROJECT_ROOT / "VERSION")

    def test_version_format(self):
        """VERSION must follow semantic versioning (X.Y.Z)."""
        self.assertRegex(
            self.version_str,
            r"^\d+\.\d+\.\d+$",
            f"VERSION '{self.version_str}' does not match semantic versioning format.",
        )

    def test_license_exists_and_has_mandatory_copyright(self):
        """LICENSE must be MIT and contain 'Copyright (c) 2026 IT Leon'."""
        license_path = PROJECT_ROOT / "LICENSE"
        self.assertTrue(license_path.is_file(), f"Missing LICENSE file at: {license_path}")

        content = license_path.read_text(encoding="utf-8")
        self.assertIn("MIT License", content)
        self.assertIn("Copyright (c) 2026 IT Leon", content)

    def test_changelog_exists_and_contains_release_version(self):
        """CHANGELOG.md must exist and contain version header matching VERSION."""
        changelog_path = PROJECT_ROOT / "CHANGELOG.md"
        self.assertTrue(changelog_path.is_file(), f"Missing CHANGELOG.md file at: {changelog_path}")

        content = changelog_path.read_text(encoding="utf-8")
        expected_version_header = f"[{self.version_str}]"
        self.assertIn(
            expected_version_header,
            content,
            f"CHANGELOG.md does not contain version section {expected_version_header}",
        )
        self.assertIn("### Added", content)
        self.assertIn("BWPVND", content)
        self.assertIn("Install.vbs", content)

    def test_readme_exists_and_contains_required_sections(self):
        """README.md must contain required copyright, UDF reference, installer and build docs."""
        readme_path = PROJECT_ROOT / "README.md"
        self.assertTrue(readme_path.is_file(), f"Missing README.md file at: {readme_path}")

        content = readme_path.read_text(encoding="utf-8")
        self.assertIn("Copyright (c) 2026 IT Leon", content)
        self.assertIn("BWPVND", content)
        self.assertIn("BWPVNDUPPER", content)
        self.assertIn("BWPVNWORDS", content)
        self.assertIn("Install.vbs", content)
        self.assertIn("build.py --release", content)
        self.assertIn("Offline", content)

    def test_buildinfo_bas_synchronized_with_version(self):
        """src/excel/BuildInfo.bas must contain current APP_VERSION and Copyright."""
        build_info_path = PROJECT_ROOT / "src" / "excel" / "BuildInfo.bas"
        self.assertTrue(build_info_path.is_file())
        content = build_info_path.read_text(encoding="utf-8")
        self.assertIn(f'Public Const APP_VERSION As String = "{self.version_str}"', content)
        self.assertIn("IT Leon", content)


class TestInstallerContract(unittest.TestCase):
    """
    Validates installer/Install.vbs syntax, structure, and functional requirements.
    """

    def setUp(self):
        self.installer_path = PROJECT_ROOT / "installer" / "Install.vbs"

    def test_installer_file_exists(self):
        """installer/Install.vbs must exist."""
        self.assertTrue(self.installer_path.is_file(), f"Missing: {self.installer_path}")

    def test_installer_contains_mandatory_contracts(self):
        """Install.vbs must contain WMI Excel detection, SHA-256 verify, MOTW removal, and COM registration."""
        code = self.installer_path.read_text(encoding="utf-8")

        # 1. WMI process detection
        self.assertIn("winmgmts:", code)
        self.assertIn("EXCEL.EXE", code)
        self.assertIn("Phat hien Microsoft Excel dang chay.", code)

        # 2. SHA-256 verification
        self.assertIn("sha256", code.lower())
        self.assertIn("Get-FileHash", code)
        self.assertIn("SHA-256 Checksum Mismatch", code)

        # 3. MOTW removal
        self.assertIn("Unblock-File", code)

        # 4. Idempotent Excel registration
        self.assertIn("Excel.Application", code)
        self.assertIn("AddIns", code)
        self.assertIn("Installed = True", code)

        # 5. Success confirmation
        self.assertIn("Cai dat BWPConvertTTNVN thanh cong!", code)

    def test_installer_syntax_execution_quiet_mode(self):
        """Install.vbs must run without syntax errors under cscript in /quiet mode."""
        cmd = ["cscript", "//nologo", str(self.installer_path), "/quiet"]
        proc = subprocess.run(cmd, capture_output=True, text=True, check=False)
        # Exit code 0 (success) or 1 (Excel is running in background) indicates valid syntax
        self.assertIn(
            proc.returncode,
            [0, 1],
            f"Install.vbs failed with syntax/runtime error (code {proc.returncode}):\n{proc.stderr}",
        )


class TestReleasePackagingPipeline(unittest.TestCase):
    """
    Validates Stage 6 release packaging, ZIP bundle completeness, and SHA-256 checksums.
    """

    @classmethod
    def setUpClass(cls):
        cls.version_str = read_version(PROJECT_ROOT / "VERSION")
        cls.dist_dir = PROJECT_ROOT / "dist"
        cls.xlam_path = PROJECT_ROOT / "dist" / "BWPConvertTTNVN.xlam"

        if not cls.xlam_path.is_file():
            raise unittest.SkipTest("BWPConvertTTNVN.xlam must exist in dist/ before release packaging tests.")

        # Run Stage 6 packaging
        cls.pkg_result = stage_6_release_packaging(PROJECT_ROOT, cls.xlam_path, cls.version_str)

    def test_sha256_checksum_files_generated(self):
        """Both versioned and canonical .sha256 checksum files must be generated."""
        versioned_sha = self.dist_dir / f"{APP_NAME}-v{self.version_str}.sha256"
        canonical_sha = self.dist_dir / f"{APP_NAME}.sha256"

        self.assertTrue(versioned_sha.is_file(), f"Missing versioned sha256 file: {versioned_sha}")
        self.assertTrue(canonical_sha.is_file(), f"Missing canonical sha256 file: {canonical_sha}")

        # Check content format
        content_versioned = versioned_sha.read_text(encoding="utf-8").strip()
        content_canonical = canonical_sha.read_text(encoding="utf-8").strip()

        expected_hash = hashlib.sha256(self.xlam_path.read_bytes()).hexdigest()
        self.assertTrue(content_versioned.startswith(expected_hash))
        self.assertTrue(content_canonical.startswith(expected_hash))
        self.assertIn("BWPConvertTTNVN.xlam", content_versioned)

    def test_zip_bundle_archive_generated_and_valid(self):
        """Release ZIP bundle must exist, be uncorrupted, and contain all 6 required items."""
        zip_path = self.dist_dir / f"{APP_NAME}-v{self.version_str}.zip"
        self.assertTrue(zip_path.is_file(), f"Missing release ZIP at: {zip_path}")
        self.assertGreater(zip_path.stat().st_size, 1000)

        with zipfile.ZipFile(zip_path, "r") as zf:
            # 1. Structural integrity check
            corrupt_file = zf.testzip()
            self.assertIsNone(corrupt_file, f"Corrupted file inside release zip: {corrupt_file}")

            # 2. Check all 6 files are present
            names = set(zf.namelist())
            expected_names = {
                "BWPConvertTTNVN.xlam",
                "Install.vbs",
                "README.md",
                "LICENSE",
                "CHANGELOG.md",
                f"{APP_NAME}-v{self.version_str}.sha256",
            }
            self.assertEqual(names, expected_names, f"ZIP contents mismatch. Found: {names}, Expected: {expected_names}")

            # 3. Check SHA-256 of .xlam inside ZIP matches actual .xlam
            xlam_in_zip_bytes = zf.read("BWPConvertTTNVN.xlam")
            actual_xlam_bytes = self.xlam_path.read_bytes()
            self.assertEqual(
                hashlib.sha256(xlam_in_zip_bytes).hexdigest(),
                hashlib.sha256(actual_xlam_bytes).hexdigest(),
                "SHA-256 of BWPConvertTTNVN.xlam inside ZIP does not match actual disk file.",
            )

    def test_build_scripts_cli_release_parameter(self):
        """scripts/build.py and scripts/build.ps1 must support --release and -Release."""
        # 1. build.py argument parser check
        import scripts.build as build_module
        import argparse

        # Verify build.py accepts --release
        proc = subprocess.run(
            [sys.executable, str(PROJECT_ROOT / "scripts" / "build.py"), "--help"],
            capture_output=True,
            text=True,
            check=True,
        )
        self.assertIn("--release", proc.stdout)

        # 2. build.ps1 parameter check
        build_ps1 = (PROJECT_ROOT / "scripts" / "build.ps1").read_text(encoding="utf-8")
        self.assertIn("[switch]$Release", build_ps1)
        self.assertIn("--release", build_ps1)


if __name__ == "__main__":
    unittest.main()
