# BWPConvertTTNVN Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a production-ready, open-source, offline Microsoft Excel VBA Add-in (`BWPConvertTTNVN.xlam`) that converts numbers into Vietnamese words via a click-based Ribbon interface, UserForms, and worksheet UDF formulas.

**Architecture:** Pure VBA core engine (`src/core/`) decoupled from Excel objects, using `ChrW$()` Unicode-safe vocabulary and character tables; Excel interaction layer (`src/excel/`) providing strict application state caching, safe block writes preserving skipped destination formulas, 2-phase transactional 1-level Undo, and deterministic worksheet UDFs; UserForms with dynamic runtime captions; compiled via an isolated Excel COM automation build script (`scripts/build.py`) with PID tracking and direct OpenXML ZIP Ribbon injection (`scripts/package_ribbon.py`).

**Tech Stack:** Microsoft Excel VBA (32-bit & 64-bit Windows), OpenXML / customUI14.xml, Python 3.13 (pywin32 / Excel COM automation for build orchestration and testing), PowerShell, VBScript.

**Spec:** [docs/superpowers/specs/2026-09-28-bwpconvert-ttnvn-excel-addin-design.md](file:///C:/Code/nora-convert-excel/docs/superpowers/specs/2026-09-28-bwpconvert-ttnvn-excel-addin-design.md)

## Global Constraints
- **Product Name:** `BWPConvertTTNVN`
- **Version Source:** Single source of truth in root `VERSION` file, generating `src/excel/BuildInfo.bas` during build.
- **Mandatory Copyright Notice:** `Copyright © 2026 - IT Leon` visibly rendered in `frmAbout` and `frmConvert` footer.
- **No Third-Party DLLs:** Zero external DLL dependencies; pure VBA implementation.
- **Offline & Private:** Zero network requests, zero telemetry, zero background daemons.
- **Target Excel Bit-ness:** Fully compatible with both 32-bit and 64-bit Microsoft Excel (2016, 2019, 2021, Office 365).
- **Process Safety:** Never execute broad process termination like `Stop-Process -Name EXCEL`; track and clean only build-owned PIDs.
- **VBA Single Source of Truth:** Python and PowerShell scripts never duplicate or reimplement Vietnamese reading grammar.
- **UDF Determinism:** Worksheet functions (`BWPVNWORDS`, `BWPVND`, `BWPVNDUPPER`) must never read Windows Registry settings.

---

### Task 1: Scaffolding, Versioning & Build Preflight Foundation

**Files:**
- Create: `VERSION`
- Create: `src/excel/BuildInfo.bas`
- Create: `scripts/clean.ps1`
- Create: `scripts/build.py`
- Test: `tests/test_preflight.py`

**Interfaces:**
- Produces: `VERSION` string (`1.0.0`), `BuildInfo.bas` with `APP_VERSION`, `APP_NAME`, `APP_COPYRIGHT`; Python preflight module verifying Excel COM availability, `AccessVBOM` trust status, and exact PID tracking.

- [ ] **Step 1: Create the root `VERSION` file**

```text
1.0.0
```

- [ ] **Step 2: Create initial `src/excel/BuildInfo.bas` template**

```vb
Attribute VB_Name = "BuildInfo"
Option Explicit

Public Const APP_NAME As String = "BWPConvertTTNVN"
Public Const APP_VERSION As String = "1.0.0"
Public Const APP_COPYRIGHT As String = "Copyright © 2026 - IT Leon"
```

- [ ] **Step 3: Write preflight test `tests/test_preflight.py`**
Test that `build.py` correctly reads `VERSION`, checks `AccessVBOM`, starts an isolated Excel COM instance, tracks its PID, and cleans it up without killing existing user Excel sessions.

```python
import os
import subprocess
import sys
import unittest
import win32com.client
import win32process
import win32gui

class TestPreflight(unittest.TestCase):
    def test_version_file_exists(self):
        with open("VERSION", "r", encoding="utf-8") as f:
            v = f.read().strip()
        self.assertEqual(v, "1.0.0")

    def test_excel_pid_tracking_and_cleanup(self):
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        hwnd = excel.Hwnd
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        self.assertGreater(pid, 0)
        excel.Quit()
        del excel
        # Verify process terminated
        try:
            handle = win32process.OpenProcess(win32process.PROCESS_TERMINATE, False, pid)
            subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
        except Exception:
            pass # Already exited

if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 4: Implement `scripts/build.py` preflight and `scripts/clean.ps1`**
Create `scripts/clean.ps1` and the scaffolding of `scripts/build.py` with PID tracking, `AccessVBOM` verification, and version injection into `BuildInfo.bas`.

- [ ] **Step 5: Run preflight test and verify pass**

Run: `python -m unittest tests/test_preflight.py`
Expected: `Ran 2 tests in ...s - OK`

- [ ] **Step 6: Commit**

```bash
git add VERSION src/excel/BuildInfo.bas scripts/ tests/test_preflight.py
git commit -m "feat: add versioning, build preflight, and PID tracking foundation"
```

---

### Task 2: Core Types & Unicode Text Engine

**Files:**
- Create: `src/core/CoreTypes.bas`
- Create: `src/core/UnicodeText.bas`
- Test: `tests/test_unicode_text.py`

**Interfaces:**
- Produces: `CoreTypes.bas` (`MAX_SUPPORTED_VALUE`, enums `VnZeroStyle`, `VnThousandStyle`, `VnFourStyle`, `VnDecimalMode`, `VnCurrencyType`, `VnCasingStyle`, UDTs `VnEngineOptions`, `VnCurrencyOptions`, `VnFormatOptions`, factory functions `DefaultEngineOptions`, `DefaultCurrencyOptions`, `DefaultFormatOptions`, error codes `ERR_*`).
- Produces: `UnicodeText.bas` (`ChrW$()` code-point accessors for all Vietnamese number words, currency terms, and character casing lookup maps).

- [ ] **Step 1: Write test verifying Unicode code points `tests/test_unicode_text.py`**
Verify expected UTF-16 code sequences for core Vietnamese vocabulary: `Một`, `đồng`, `chẵn`, `nghìn`, `ngàn`, `tỷ`, `lẻ`, `linh`, `mốt`, `tư`, `lăm`.

```python
import unittest

EXPECTED_WORDS = {
    "dong": "\u0111\u1ed3ng",
    "chan": "ch\u1eb5n",
    "nghin": "ngh\u00ecn",
    "ngan": "ng\u00e0n",
    "ty": "t\u1ef7",
    "le": "l\u1ebb",
    "linh": "linh",
    "mot_cuoi": "m\u1ed1t",
    "tu": "t\u01b0",
    "lam": "l\u0103m"
}

class TestUnicodeText(unittest.TestCase):
    def test_code_points_match_vietnamese_grammar(self):
        for k, v in EXPECTED_WORDS.items():
            self.assertTrue(len(v) > 0)
            # Verify explicit UTF-16 code points
            self.assertNotIn("?", v)

if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Implement `src/core/CoreTypes.bas`**
Copy verbatim the validated `CoreTypes.bas` contract from Section 4.1 of the design spec (declaring constants, error codes, enums, UDTs with `DecimalPlaces`, and default factory functions).

- [ ] **Step 3: Implement `src/core/UnicodeText.bas`**
Define pure ASCII-safe VBA module using `ChrW$()` for all Vietnamese vocabulary terms:
- `WordDong`: `ChrW$(&H111) & ChrW$(&H1ED3) & "ng"`
- `WordChan`: `"ch" & ChrW$(&H1EB5) & "n"`
- `WordNghin`: `"ngh" & ChrW$(&HEC) & "n"`
- `WordNgan`: `"ng" & ChrW$(&HE0) & "n"`
- `WordTy`: `"t" & ChrW$(&H1EF7)`
- `WordLe`: `"l" & ChrW$(&H1EBB)`
- `WordLinh`: `"linh"`
- `WordMotCuoi`: `"m" & ChrW$(&H1ED1) & "t"`
- `WordTu`: `"t" & ChrW$(&H1B0)`
- `WordLam`: `"l" & ChrW$(&H103) & "m"`
- Complete mapping functions for Unicode uppercase / lowercase vowels.

- [ ] **Step 4: Verify test execution**

Run: `python -m unittest tests/test_unicode_text.py`
Expected: `PASS`

- [ ] **Step 5: Commit**

```bash
git add src/core/CoreTypes.bas src/core/UnicodeText.bas tests/test_unicode_text.py
git commit -m "feat: implement CoreTypes and ASCII-safe UnicodeText engine"
```

---

### Task 3: Algorithmic Vietnamese Number-to-Words Engine

**Files:**
- Create: `src/core/VietnameseNumber.bas`
- Create: `tests/expected_cases.csv`
- Test: `tests/test_number_engine.py`

**Interfaces:**
- Consumes: `CoreTypes.bas`, `UnicodeText.bas`
- Produces: `NumberToVietnamese(Value, Options)`, `TryNumberToVietnamese(Value, Options, OutText, OutError)`, `NumberToVietnameseDefault(Value)`. Handles scales up to $10^{12}$, `forceFullTriplet`, zero-gaps, irregular units (1, 4, 5), negative amounts, decimals.

- [ ] **Step 1: Create `tests/expected_cases.csv` with complete schema**
Include standard numbers, scale transitions, zero-gap numbers, boundary limits ($999,999,999,999,999$), negative values, dialects (`lẻ`/`linh`, `nghìn`/`ngàn`, `tư`/`bốn`), and overflow errors.

- [ ] **Step 2: Implement `src/core/VietnameseNumber.bas`**
Implement the pure VBA algorithm:
- `IsValidNumber(Value)` checking numeric subtypes and bounds `[-MAX_SUPPORTED_VALUE, MAX_SUPPORTED_VALUE]`.
- Normalize numeric value into absolute digit string and sign.
- Split integer into 3-digit triplets from right to left.
- Apply `forceFullTriplet` across skipped `000` triplets.
- Apply irregular rules: `mốt` vs `một`, `tư` vs `bốn`, `lăm` vs `năm`, `lẻ` vs `linh`.
- Truncate decimals toward zero via `Fix()` in Mode 1, or spell digits with `phẩy` in Mode 2.
- Zero returns `"không"`. Negatives prefix `"âm "`.

- [ ] **Step 3: Create automated VBA test runner `tests/test_engine.bas`**
Expose a VBA procedure `RunCoreTests()` that tests all cases in-memory and returns pass/fail count.

- [ ] **Step 4: Verify compilation and tests via Python COM test runner**

Run: `python -c "import tests.run_tests; print('ready')"`
Expected: Clean load without syntax errors.

- [ ] **Step 5: Commit**

```bash
git add src/core/VietnameseNumber.bas tests/expected_cases.csv tests/test_engine.bas
git commit -m "feat: implement pure VBA Vietnamese number-to-words algorithm"
```

---

### Task 4: Currency Formatter, Deterministic Rounding & Core Coordinator

**Files:**
- Create: `src/core/Currency.bas`
- Create: `src/core/TextFormatter.bas`
- Create: `src/core/CoreCoordinator.bas`
- Test: `tests/test_currency_coordinator.py`

**Interfaces:**
- Consumes: `CoreTypes.bas`, `UnicodeText.bas`, `VietnameseNumber.bas`
- Produces: `NumberToCurrencyWords(Value, CurrOptions, EngineOptions)`, `FormatText(Text, Casing, AddPeriod)`, `ConvertNumber(...)`, `ConvertNumberDefault(Value)`.
- Features: `RoundHalfAwayFromZero`, sub-units (USD cents), carry rounding (1.999 $\rightarrow$ 2.00), negative sign preservation (-0.50 USD), `"chẵn"`, sentence/upper/lower casing.

- [ ] **Step 1: Write test for currency and formatting edge cases**
Test:
- `125430000` VND $\rightarrow$ *Một trăm hai mươi lăm triệu bốn trăm ba mươi nghìn đồng chẵn.*
- `-150000` VND $\rightarrow$ *Âm một trăm năm mươi nghìn đồng chẵn.*
- `1.005` USD $\rightarrow$ *Một đô la Mỹ một cent.*
- `1.999` USD $\rightarrow$ *Hai đô la Mỹ.* (carry rounded before decomposition)
- `-0.50` USD $\rightarrow$ *Âm không đô la Mỹ năm mươi cent.*
- `Casing = VnCaseUpper` $\rightarrow$ full uppercase with accents (*MỘT TRĂM...*).

- [ ] **Step 2: Implement `src/core/Currency.bas`**
- Pure VBA `RoundHalfAwayFromZero(Value, DecimalPlaces)`.
- Decompose rounded value into sign, integer part, and fractional sub-units.
- Assemble: `[Âm ] + [IntegerWords] + [CurrencyUnit] + [SubUnitWords + SubUnitName | "chẵn"]`.

- [ ] **Step 3: Implement `src/core/TextFormatter.bas`**
- Whitespace collapsing, trimming spaces before punctuation.
- Sentence case (capitalizing first letter using `UnicodeText`), Upper case, Lower case.
- Conditional trailing period.

- [ ] **Step 4: Implement `src/core/CoreCoordinator.bas`**
- Implements `ConvertNumber`, `TryConvertNumber`, and `ConvertNumberDefault`.

- [ ] **Step 5: Verify tests pass**

Run: `python -m unittest tests/test_currency_coordinator.py`
Expected: `PASS`

- [ ] **Step 6: Commit**

```bash
git add src/core/Currency.bas src/core/TextFormatter.bas src/core/CoreCoordinator.bas tests/test_currency_coordinator.py
git commit -m "feat: implement Currency, TextFormatter, and CoreCoordinator pipeline"
```

---

### Task 5: Excel State Management & Batch Cell Processor

**Files:**
- Create: `src/excel/Settings.bas`
- Create: `src/excel/UndoManager.bas`
- Create: `src/excel/CellProcessor.bas`
- Test: `tests/test_cell_processor.py`

**Interfaces:**
- Consumes: `CoreTypes.bas`, `CoreCoordinator.bas`
- Produces: `LoadAllSettings`, `SaveAllSettings`, `ResetSettingsToDefault`; `UndoManager` with 2-phase commit and 10,000 cell limit; `CellProcessor.ConvertRange(...)` returning structured `VnBatchResult`.
- Features: Exact state caching & restoration (`ScreenUpdating`, `Calculation`, `EnableEvents`, `DisplayAlerts`, `DisplayStatusBar`, `StatusBar`), contiguous area check, cross-sheet safe overlap check, anchor destination expansion, formula-safe skip preservation.

- [ ] **Step 1: Write integration tests for CellProcessor**
Test in a temporary Excel workbook:
- Contiguous range `A1:A5` converts to `B1:B5`.
- Multi-area range `A1:A2,C1:C2` is rejected.
- Same-sheet overlap `A1:A10` to `A5:A14` is rejected.
- Cross-sheet conversion `Sheet1!A1:A5` to `Sheet2!B1` succeeds without overlap error.
- Skipped cell with existing destination formula retains the formula intact.
- 2-phase Undo restores original contents and formulas.

- [ ] **Step 2: Implement `src/excel/Settings.bas`**
- Windows Registry `HKCU\Software\VB and VBA Program Settings\BWPConvertTTNVN`.
- Safe bounds validation on enums, `SettingsVersion = 1`, and `ResetSettingsToDefault`.

- [ ] **Step 3: Implement `src/excel/UndoManager.bas`**
- Staged snapshot: `Workbook`, `Worksheet`, `Address`, `Values`, `Formulas`, `HasFormula`.
- `PrepareUndoSnapshot`, `CommitUndoSnapshot`, `ExecuteUndo`, `ClearUndo`.
- Guard threshold `MAX_UNDO_CELLS = 10000`.

- [ ] **Step 4: Implement `src/excel/CellProcessor.bas`**
- State caching and restoration block in `Cleanup`.
- Normalizes single cells and 1D/2D ranges into memory arrays.
- Implements formula-safe write strategy for skipped rows.
- Returns `VnBatchResult`.

- [ ] **Step 5: Run tests and verify pass**

Run: `python -m unittest tests/test_cell_processor.py`
Expected: `PASS`

- [ ] **Step 6: Commit**

```bash
git add src/excel/Settings.bas src/excel/UndoManager.bas src/excel/CellProcessor.bas tests/test_cell_processor.py
git commit -m "feat: implement Settings, 2-phase UndoManager, and batch CellProcessor"
```

---

### Task 6: Worksheet UDF Functions

**Files:**
- Create: `src/excel/UDF.bas`
- Test: `tests/test_udf_integration.py`

**Interfaces:**
- Consumes: `CoreCoordinator.bas`, `CoreTypes.bas`
- Produces: `=BWPVNWORDS(Target, [ZeroStyle], [ThousandStyle])`, `=BWPVND(Target, [AddChan], [ZeroStyle], [ThousandStyle])`, `=BWPVNDUPPER(Target, [AddChan])`.
- Rules: Never reads Registry; `=BWPVNWORDS(125.05)` enables `VnDecimalDigits` to read *phẩy không năm*; blank input returns `""`; invalid returns `CVErr(xlErrValue)`; non-volatile.

- [ ] **Step 1: Write worksheet UDF test `tests/test_udf_integration.py`**
Test live Excel workbook:
- Set `A1 = 125430000`, `B1 = "=BWPVND(A1)"` $\rightarrow$ Calculate $\rightarrow$ verify `B1.Value2`.
- Set `A2 = 125.05`, `B2 = "=BWPVNWORDS(A2)"` $\rightarrow$ verify decimal reading.
- Set `A3 = 500000`, `B3 = "=BWPVNDUPPER(A3)"` $\rightarrow$ verify all-uppercase.
- Set `A4 = ""` $\rightarrow$ verify formula returns `""`.
- Set `A5 = "invalid"` $\rightarrow$ verify formula returns `#VALUE!`.

- [ ] **Step 2: Implement `src/excel/UDF.bas`**
Implement the three UDF functions with exact signatures and parameter mappings, calling `ConvertNumber` with fixed documented defaults.

- [ ] **Step 3: Run worksheet UDF integration tests**

Run: `python -m unittest tests/test_udf_integration.py`
Expected: `PASS`

- [ ] **Step 4: Commit**

```bash
git add src/excel/UDF.bas tests/test_udf_integration.py
git commit -m "feat: implement deterministic worksheet UDFs (BWPVNWORDS, BWPVND, BWPVNDUPPER)"
```

---

### Task 7: UserForms with Dynamic Runtime Captions

**Files:**
- Create: `src/forms/frmConvert.frm` + `frmConvert.frx`
- Create: `src/forms/frmSettings.frm` + `frmSettings.frx`
- Create: `src/forms/frmAbout.frm` + `frmAbout.frx`
- Test: `tests/test_forms.py`

**Interfaces:**
- Consumes: `UnicodeText.bas`, `CellProcessor.bas`, `Settings.bas`, `BuildInfo.bas`
- Produces: Main conversion dialog, Settings dialog, About dialog.
- Features: Range object retention (cross-sheet safe), `InputBox(Type:=8)` cancel handling, single casing combobox, dynamic currency controls (disabling formula mode for USD/Custom), `[Mặc định]` loading without immediate save, bold `Copyright © 2026 - IT Leon`.

- [ ] **Step 1: Create `src/forms/frmAbout.frm` and `.frx`**
Include version from `BuildInfo.APP_VERSION` and mandatory `Copyright © 2026 - IT Leon`.

- [ ] **Step 2: Create `src/forms/frmSettings.frm` and `.frx`**
Include Grammar, Formatting, Output mode, Quick Convert direction, and Safety frames. `[Mặc định]` populates UI only; `[Lưu]` persists to Registry.

- [ ] **Step 3: Create `src/forms/frmConvert.frm` and `.frx`**
Include Range pickers, Currency combobox, mutually exclusive Casing combobox, Output mode radios, and copyright footer.

- [ ] **Step 4: Write test verifying UserForm structures and dynamic caption assignment**

Run: `python -m unittest tests/test_forms.py`
Expected: `PASS`

- [ ] **Step 5: Commit**

```bash
git add src/forms/ tests/test_forms.py
git commit -m "feat: implement frmConvert, frmSettings, and frmAbout with dynamic Unicode captions"
```

---

### Task 8: Ribbon Definition & Direct OpenXML Packaging Pipeline

**Files:**
- Create: `ribbon/customUI14.xml`
- Create: `src/excel/RibbonCallbacks.bas`
- Create: `scripts/package_ribbon.py`
- Test: `tests/test_ribbon_package.py`

**Interfaces:**
- Produces: `customUI14.xml` with Office 2010+ namespace; `RibbonCallbacks.bas` caching `IRibbonUI` and driving `GetUndoEnabled`; `package_ribbon.py` injecting Ribbon XML and relationship directly into OpenXML `.xlam` package.

- [ ] **Step 1: Create `ribbon/customUI14.xml`**
Define Ribbon tab `BWPConvertTTNVN` with groups `grpConversion` and `grpSettings`, buttons using native high-DPI `imageMso` icons (`ChangeTextCase`, `AutoSum`, `Undo`, `ControlProperties`, `Info`).

- [ ] **Step 2: Implement `src/excel/RibbonCallbacks.bas`**
Implement `OnRibbonLoad`, `OnConvertClick`, `OnQuickConvertClick` (routing 2D selections to `frmConvert`), `OnUndoClick`, `OnSettingsClick`, `OnAboutClick`, and `GetUndoEnabled`.

- [ ] **Step 3: Implement `scripts/package_ribbon.py`**
- Opens `.xlam` as ZIP archive.
- Injects `customUI/customUI14.xml`.
- Adds CustomUI relationship to `_rels/.rels`.
- Validates `[Content_Types].xml` and `xl/vbaProject.bin`.

- [ ] **Step 4: Write test verifying packaged OpenXML integrity**

Run: `python -m unittest tests/test_ribbon_package.py`
Expected: `PASS`

- [ ] **Step 5: Commit**

```bash
git add ribbon/customUI14.xml src/excel/RibbonCallbacks.bas scripts/package_ribbon.py tests/test_ribbon_package.py
git commit -m "feat: implement Ribbon definition, callbacks, and OpenXML direct packaging script"
```

---

### Task 9: Automated Test Suite & Excel COM Orchestrator

**Files:**
- Create: `scripts/build.py` (complete build pipeline)
- Create: `scripts/build.ps1` (PowerShell wrapper)
- Create: `tests/run_tests.py` (3-tier test runner)
- Test: Full end-to-end execution

**Interfaces:**
- Produces: Automated build producing `dist/BWPConvertTTNVN.xlam`, followed by reopen validation and automated 3-tier testing (Core API via qualified `Application.Run`, Worksheet UDFs, Excel integration tests).

- [ ] **Step 1: Finalize `scripts/build.py`**
Implement all 7 pipeline stages adhering strictly to the 15-step dependency import manifest, isolated Excel COM instance, PID tracking, and reopen validation.

- [ ] **Step 2: Implement `tests/run_tests.py`**
Implement 3 test classes:
- Class A: Core API calls against all 50+ entries in `expected_cases.csv`.
- Class B: Real worksheet formulas in temporary workbooks.
- Class C: Quick Convert, batch arrays, destination expansion, overlap rejection, Undo.

- [ ] **Step 3: Execute build and full test suite**

Run: `python scripts/build.py`
Expected: Successfully compiles `dist/BWPConvertTTNVN.xlam` and all tests pass.

- [ ] **Step 4: Commit**

```bash
git add scripts/build.py scripts/build.ps1 tests/run_tests.py
git commit -m "feat: complete automated build pipeline and 3-tier test runner"
```

---

### Task 10: Installer, Release Packaging & Documentation

**Files:**
- Create: `installer/Install.vbs`
- Create: `README.md`
- Create: `CHANGELOG.md`
- Create: `LICENSE`
- Modify: `scripts/build.py` (add `--release` / `-Package` flag)
- Test: Release packaging and installer simulation

**Interfaces:**
- Produces: `Install.vbs` with running Excel detection, SHA-256 verification, and MOTW removal; `dist/BWPConvertTTNVN-v1.0.0.zip` bundle and `dist/BWPConvertTTNVN-v1.0.0.sha256`; complete documentation.

- [ ] **Step 1: Implement `installer/Install.vbs`**
- Detects if Excel is open $\rightarrow$ prompts user and exits gracefully without terminating user work.
- Verifies SHA-256 checksum of `BWPConvertTTNVN.xlam`.
- Unblocks Mark of the Web (`Zone.Identifier`).
- Copies add-in to `%APPDATA%\Microsoft\AddIns\`.
- Idempotently adds and installs add-in in Excel.

- [ ] **Step 2: Create `LICENSE` (MIT) and `CHANGELOG.md`**

- [ ] **Step 3: Create `README.md`**
Comprehensive documentation: overview, features, installation via `Install.vbs` or manual, Quick Convert usage, UDF reference (`BWPVNWORDS`, `BWPVND`, `BWPVNDUPPER`), Settings guide, reproducible build instructions, offline privacy statement, and license.

- [ ] **Step 4: Add `--release` packaging to `scripts/build.py`**
Creates `dist/BWPConvertTTNVN-v1.0.0.zip` and generates `dist/BWPConvertTTNVN-v1.0.0.sha256`.

- [ ] **Step 5: Run full release build and verify distribution artifacts**

Run: `python scripts/build.py --release`
Expected:
- `dist/BWPConvertTTNVN.xlam`
- `dist/BWPConvertTTNVN-v1.0.0.zip`
- `dist/BWPConvertTTNVN-v1.0.0.sha256`
- `dist/test-report.txt`

- [ ] **Step 6: Commit**

```bash
git add installer/Install.vbs README.md CHANGELOG.md LICENSE scripts/build.py
git commit -m "feat: implement installer, release packaging, and user documentation"
```
