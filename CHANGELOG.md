# Changelog

All notable changes to the BWPConvertTTNVN project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-09-28

### Added
- **Vietnamese Number Conversion Core**:
  - High-precision number-to-words algorithm supporting integer and decimal values up to 15 digits (`999,999,999,999,999`).
  - Standard Vietnamese grammatical rules: *mười một*, *mười lăm*, *hai mươi mốt*, *hai mươi tư*, *hai mươi lăm*, *lẻ/linh*, *không trăm lẻ*.
  - Multi-currency support: VND (`đồng chẵn`), USD (`đô la Mỹ`, `cent`), EUR (`euro`, `cent`), JPY (`yên Nhật`), GBP (`bảng Anh`, `pence`), CNY (`nhân dân tệ`, `hào`, `xu`).
  - 100% 7-bit ASCII safe VBA source code using runtime UTF-16 `ChrW` array reconstruction to guarantee flawless diacritics across any Windows locale and system codepage.
  - Granular regional dialect and style configurations (*lẻ* vs *linh*, *nghìn* vs *ngàn*, *bốn* vs *tư*).
  - Multiple text casing options: Sentence case (Capitalize First), Title Case, Lowercase, and UPPERCASE.
- **Excel Custom Worksheet Functions (UDFs)**:
  - `=BWPVNWORDS(value, [currency], [style], [unit])`: Full-featured conversion function with customizable parameters.
  - `=BWPVND(value)`: High-performance canonical shorthand for Vietnamese Dong currency reading.
  - `=BWPVNDUPPER(value)`: Shorthand returning all-uppercase Vietnamese Dong phrase.
  - Compatibility aliases: `=VND()`, `=VNWORDS()`, `=VNDUPPER()`.
  - Excel Formula Wizard parameter metadata and descriptive tooltips via `Application.MacroOptions`.
- **Modern User Interface & UserForms**:
  - Dedicated Microsoft Excel Ribbon tab (`BWPConvertTTNVN`) with custom icons and organized action groups.
  - Interactive Quick Convert dialog (`frmConvert`) supporting live calculation preview, range picker, and destination selection.
  - User Preferences dialog (`frmSettings`) with persistent Windows Registry configuration (`HKCU\Software\BWPConvertTTNVN`).
  - About dialog (`frmAbout`) presenting version info, system diagnostics, and copyright metadata.
- **Batch Processing & Transactional Undo**:
  - High-performance 2D variant array chunk processing capable of converting 10,000 cells in under 1.5 seconds.
  - Destination overlap protection preventing accidental data overwrites on the same worksheet.
  - Safe formula preservation skipping empty or invalid cells without destroying existing destination formulas.
  - Full transactional single-level Undo restoring original values, formulas, and clearing inserted entries.
- **Automated Installer & Release Packaging**:
  - Double-clickable, non-elevated VBScript installer (`Install.vbs`) running under standard user permissions.
  - Running Excel process detection prompting the user to save work without process termination.
  - Cryptographic SHA-256 integrity verification validating package authenticity before installation.
  - Mark of the Web (MOTW / `Zone.Identifier`) removal unblocking downloaded add-in files.
  - Idempotent add-in registration in `%APPDATA%\Microsoft\AddIns` with automated Excel COM integration.
  - Automated release packaging generating standardized ZIP distribution (`BWPConvertTTNVN-v1.0.0.zip`) and SHA-256 checksums (`BWPConvertTTNVN-v1.0.0.sha256`).
- **Reproducible Build Pipeline & Automated 3-Tier Test Suite**:
  - Isolated Excel COM build pipeline with strict PID tracking and process cleanup.
  - Direct OpenXML Ribbon packaging injecting `customUI14.xml` and updating relationship manifests.
  - Comprehensive 3-tier automated test suite covering core algorithms, worksheet UDFs, and Excel COM integration.
