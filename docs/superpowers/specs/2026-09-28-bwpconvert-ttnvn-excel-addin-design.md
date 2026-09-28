# BWPConvertTTNVN — Technical Design Specification
**Vietnamese Number-to-Words Excel Add-in (.xlam)**  
**Version:** 1.0.0  
**Date:** 2026-09-28  
**Author/Copyright:** Copyright © 2026 - IT Leon  
**License:** MIT License  

---

## 1. Executive Summary & Goals

### 1.1 Product Purpose
**BWPConvertTTNVN** is a lightweight, open-source, fully offline Microsoft Excel Add-in (`.xlam`) designed to convert numeric values into Vietnamese words. Built primarily for Vietnamese accountants, finance, administrative, and sales personnel, it provides a native, modern, click-based workflow inspired by traditional tools like VNTools, but implemented from scratch with clean, auditable VBA, reproducible OpenXML packaging, and zero external binary dependencies.

### 1.2 Core Tenets
1. **SELECT → CLICK → CONVERT**: Zero formula typing required for standard usage.
2. **Deterministic & Locale-Safe**: Strict separation between pure mathematical number grammar and Excel UI/formatting.
3. **100% Offline & Private**: Zero network access, zero external telemetry, zero bundled third-party DLLs.
4. **VBA Engine as Single Source of Truth**: External tooling (Python/PowerShell) exists solely for build orchestration and test automation via Excel COM, never duplicating conversion logic.
5. **Excel Application State Safety**: Flawless caching and restoration of Excel runtime states (`ScreenUpdating`, `Calculation`, `EnableEvents`, `DisplayAlerts`, `StatusBar`) across all execution and error paths.
6. **Mandatory Interface Notice**: Product UI visibly displays `Copyright © 2026 - IT Leon` in the About dialog (`frmAbout`) and main conversion footer (`frmConvert`).

---

## 2. Project Directory Structure

```text
BWPConvertTTNVN/
│
├── VERSION                               # Single source of truth for version (e.g. "1.0.0")
├── LICENSE                               # MIT License
├── README.md                             # Documentation & user guide
├── CHANGELOG.md                          # Release history
│
├── src/
│   ├── core/                             # Pure VBA: zero Excel dependencies, zero UI
│   │   ├── CoreTypes.bas                 # Constants, Enums, UDTs, error codes, option factories
│   │   ├── VietnameseNumber.bas          # Algorithmic number-to-words reading engine
│   │   ├── Currency.bas                  # Currency composition, sub-units, "chẵn" handling
│   │   ├── TextFormatter.bas             # Pure VBA Unicode casing, whitespace, punctuation
│   │   └── CoreCoordinator.bas           # Unified entry point pipeline (Simple & Advanced APIs)
│   │
│   ├── excel/                            # Excel object model integration
│   │   ├── CellProcessor.bas             # Range & batch engine, state cache, bulk arrays
│   │   ├── UDF.bas                       # Worksheet functions (=BWPVNWORDS, =BWPVND, =BWPVNDUPPER)
│   │   ├── Settings.bas                  # Registry persistence (HKCU) with versioning & reset
│   │   ├── UndoManager.bas               # 2-phase 1-level transactional Undo (max 10,000 cells)
│   │   └── RibbonCallbacks.bas           # Ribbon event handlers, state invalidator, selection guards
│   │
│   └── forms/                            # UserForms (.frm + .frx binary pairs)
│       ├── frmConvert.frm
│       ├── frmConvert.frx
│       ├── frmSettings.frm
│       ├── frmSettings.frx
│       ├── frmAbout.frm
│       └── frmAbout.frx
│
├── ribbon/
│   ├── customUI14.xml                    # Office 2010+ Ribbon XML definition
│   └── icons/                            # Optional custom assets (built-in imageMso prioritized)
│
├── tests/
│   ├── test_engine.bas                   # In-VBA automated test runner
│   ├── expected_cases.csv                # Comprehensive test cases matrix (50+ scenarios)
│   └── run_tests.py                      # Python COM orchestrator driving real Excel instance
│
├── scripts/
│   ├── build.py                          # Primary Python build script (PID tracking, COM import)
│   ├── build.ps1                         # PowerShell build wrapper with optional AccessVBOM preflight
│   ├── package_ribbon.py                 # OpenXML ZIP direct injector for customUI14.xml
│   └── clean.ps1                         # Safe cleanup targeting only build-owned PIDs
│
├── installer/
│   └── Install.vbs                       # Double-clickable no-admin installer with MOTW removal
│
├── dist/                                 # Build distribution outputs
│   ├── BWPConvertTTNVN.xlam              # Compiled Excel Add-in
│   ├── BWPConvertTTNVN-v1.0.0.zip        # Release bundle
│   ├── BWPConvertTTNVN-v1.0.0.sha256     # Cryptographic integrity checksum
│   └── test-report.txt                   # Automated test run results
│
└── docs/
    └── superpowers/specs/                # Architecture specifications
```

---

## 3. Vietnamese Number-Reading Rules & Grammar

### 3.1 Vocabulary & Grammar Toggles
* **Digits (0–9)**: `không`, `một`, `hai`, `ba`, `bốn`, `năm`, `sáu`, `bảy`, `tám`, `chín`.
* **Zero / Linking Word (`VnZeroStyle`)**:
  * `VnZeroLe = 0` (Default): `"lẻ"` (e.g. *một trăm lẻ năm*).
  * `VnZeroLinh = 1`: `"linh"` (e.g. *một trăm linh năm*).
* **Thousands Scale Word (`VnThousandStyle`)**:
  * `VnThousandNghin = 0` (Default, Northern): `"nghìn"`.
  * `VnThousandNgan = 1` (Southern): `"ngàn"`.
* **Four Style After Tens (`VnFourStyle`)**:
  * `VnFourTu = 0` (Default): `"tư"` when tens $\ge 2$ (e.g. `24` $\rightarrow$ *hai mươi tư*, `124` $\rightarrow$ *một trăm hai mươi tư*).
  * `VnFourBon = 1`: `"bốn"` (e.g. `24` $\rightarrow$ *hai mươi bốn*).
  * Note: `4` alone is always *bốn*, and `14` is always *mười bốn*.
* **Five Style (`c_unit = 5`)**:
  * `c_unit = 5` and tens $\ge 1$: `"lăm"` (e.g. `15` $\rightarrow$ *mười lăm*, `25` $\rightarrow$ *hai mươi lăm*).
  * `c_unit = 5` and tens $= 0$: `"năm"` (e.g. `5` $\rightarrow$ *năm*, `105` $\rightarrow$ *một trăm lẻ năm*).
* **One Style (`c_unit = 1`)**:
  * `c_unit = 1` and tens $\ge 2$: `"mốt"` (e.g. `21` $\rightarrow$ *hai mươi mốt*, `121` $\rightarrow$ *một trăm hai mươi mốt*).
  * `c_unit = 1` and tens $< 2$: `"một"` (e.g. `1` $\rightarrow$ *một*, `11` $\rightarrow$ *mười một*, `101` $\rightarrow$ *một trăm lẻ một*).

### 3.2 Scales & Triplet Grammar
* **Supported Scales for MVP**:
  * $10^0$: Units
  * $10^3$: `nghìn` / `ngàn`
  * $10^6$: `triệu`
  * $10^9$: `tỷ`
  * $10^{12}$: `nghìn tỷ` / `ngàn tỷ`
  * *(Scale $10^{15}$: `triệu tỷ` reserved for future range expansion).*
* **Maximum Absolute Limit**:
  * `Public Const MAX_SUPPORTED_VALUE As Double = 999999999999999#` ($999,999,999,999,999$).
  * Exceeding this boundary strictly triggers `ERR_OUT_OF_RANGE`.
* **Zero Handling**:
  * Number `0` evaluates strictly to `"không"`.
* **Negative Numbers**:
  * Values $< 0$ prefix `"âm "` followed by the reading of the absolute value.
* **Deterministic Missing-Hundreds Rule Across Zero Triplets (`forceFullTriplet`)**:
  * For any 3-digit triplet `[hundred, ten, unit]`:
    * If higher non-zero triplets exist in the number, `forceFullTriplet = True`.
    * When `forceFullTriplet = True` and `hundred = 0` (while `ten > 0` or `unit > 0`), the hundreds place MUST be spoken as `"không trăm"` (e.g. `1,000,005` $\rightarrow$ *một triệu không trăm lẻ năm*).
    * Triplet `000` in the middle of a number is omitted from speech, but triggers `forceFullTriplet = True` for subsequent non-zero triplets.
    * Triplet `000` at the end of a number produces no speech (e.g. `1,000,000` $\rightarrow$ *một triệu*).

### 3.3 Decimal Handling
* **Mode 1: Truncate toward zero (`VnDecimalIgnore`)**:
  * Uses VBA `Fix()` behavior (truncation toward zero), not mathematical floor `Int()`.
  * For positive numbers: `Fix(125.9) = 125`.
  * For negative numbers: `Fix(-125.9) = -125`.
  * Default mode for VND.
* **Mode 2: Read decimal digits (`VnDecimalDigits`)**:
  * Reads integer part, appends `"phẩy"`, then reads each decimal digit individually (e.g. `125.05` $\rightarrow$ *một trăm hai mươi lăm phẩy không năm*).
* **Mode 3: Currency sub-units**:
  * Governed by currency configuration. Fractional remainder is rounded using half-away-from-zero logic to `CustomDecimals`.

---

## 4. Core Engine Public API (`src/core/`)

All core modules are pure VBA with zero references to Excel objects (`Range`, `Application`, `Worksheet`), Windows Registry, or external DLLs.

### 4.1 `CoreTypes.bas`
Declares all global constants, enums, error codes, UDTs, and default factory functions.

```vb
Option Explicit

Public Const MAX_SUPPORTED_VALUE As Double = 999999999999999#

' Core Error Codes
Public Const ERR_INVALID_NUMBER As Long = vbObjectError + 2101
Public Const ERR_OUT_OF_RANGE As Long = vbObjectError + 2102
Public Const ERR_INVALID_OPTIONS As Long = vbObjectError + 2103

Public Enum VnZeroStyle
    VnZeroLe = 0
    VnZeroLinh = 1
End Enum

Public Enum VnThousandStyle
    VnThousandNghin = 0
    VnThousandNgan = 1
End Enum

Public Enum VnFourStyle
    VnFourTu = 0
    VnFourBon = 1
End Enum

Public Enum VnDecimalMode
    VnDecimalIgnore = 0
    VnDecimalDigits = 1
End Enum

Public Enum VnCurrencyType
    VnCurrVND = 0
    VnCurrUSD = 1
    VnCurrNone = 2
    VnCurrCustom = 3
End Enum

Public Enum VnCasingStyle
    VnCaseSentence = 0
    VnCaseUpper = 1
    VnCaseLower = 2
End Enum

Public Type VnEngineOptions
    ZeroStyle As VnZeroStyle
    ThousandStyle As VnThousandStyle
    FourStyle As VnFourStyle
    DecimalMode As VnDecimalMode
End Type

Public Type VnCurrencyOptions
    CurrencyType As VnCurrencyType
    AddChan As Boolean
    CustomPrefix As String
    CustomSuffix As String
    CustomSubUnit As String
    CustomDecimals As Integer
End Type

Public Type VnFormatOptions
    Casing As VnCasingStyle
    AddPeriod As Boolean
End Type

' Factory Functions for Sensible Defaults
Public Function DefaultEngineOptions() As VnEngineOptions
    Dim Opts As VnEngineOptions
    Opts.ZeroStyle = VnZeroLe
    Opts.ThousandStyle = VnThousandNghin
    Opts.FourStyle = VnFourTu
    Opts.DecimalMode = VnDecimalIgnore
    DefaultEngineOptions = Opts
End Function

Public Function DefaultCurrencyOptions( _
    Optional ByVal CurrencyType As VnCurrencyType = VnCurrVND _
) As VnCurrencyOptions
    Dim Opts As VnCurrencyOptions
    Opts.CurrencyType = CurrencyType
    Opts.AddChan = True
    Select Case CurrencyType
        Case VnCurrVND
            Opts.CustomDecimals = 0
        Case VnCurrUSD
            Opts.CustomDecimals = 2
        Case Else
            Opts.CustomDecimals = 0
    End Select
    DefaultCurrencyOptions = Opts
End Function

Public Function DefaultFormatOptions() As VnFormatOptions
    Dim Opts As VnFormatOptions
    Opts.Casing = VnCaseSentence
    Opts.AddPeriod = True
    DefaultFormatOptions = Opts
End Function
```

### 4.2 `VietnameseNumber.bas`
Pure number grammar conversion engine.
* **Strict Input Typing**: Accepts numeric Variant subtypes (`vbInteger`, `vbLong`, `vbSingle`, `vbDouble`, `vbCurrency`, `vbDecimal`). Rejects non-numeric strings (string coercion is owned by `src/excel/`).
* **Public Signatures**:
  ```vb
  ' Low-level strict function (raises ERR_* on failure)
  Public Function NumberToVietnamese( _
      ByVal Value As Variant, _
      ByRef Options As VnEngineOptions _
  ) As String

  ' Safe function returning Boolean status
  Public Function TryNumberToVietnamese( _
      ByVal Value As Variant, _
      ByRef Options As VnEngineOptions, _
      ByRef OutText As String, _
      ByRef OutErrorMessage As String _
  ) As Boolean

  ' Convenience wrapper with default options
  Public Function NumberToVietnameseDefault(ByVal Value As Variant) As String
  ```

### 4.3 `Currency.bas`
Handles monetary assembly, rounding, sub-units, and `"chẵn"`.
* **Deterministic Rounding**: Implements pure VBA `RoundHalfAwayFromZero(Value, DecimalPlaces)`—never uses VBA `Round()` (which performs banker's rounding).
* **Carry Handling**: Rounding executes on the original value *before* integer and sub-unit decomposition (e.g. `1.999 USD` rounds to `2.00 USD`, avoiding `1 dollar 100 cents`).
* **Sign Safety**: Inspects raw signed input once; applies `"âm "` before the entire currency phrase (e.g. `-0.50 USD` $\rightarrow$ *âm không đô la Mỹ năm mươi cent*).
* **Public Signature**:
  ```vb
  Public Function NumberToCurrencyWords( _
      ByVal Value As Variant, _
      ByRef CurrOptions As VnCurrencyOptions, _
      ByRef EngineOptions As VnEngineOptions _
  ) As String
  ```

### 4.4 `TextFormatter.bas`
Handles spacing, Unicode casing, and trailing punctuation in pure VBA.
* **Pure VBA Unicode Mapping**: Implements full lookup tables for Vietnamese accented characters (e.g. `ă/Ă`, `đ/Đ`, `ê/Ê`, `ơ/Ơ`, `ư/Ư`, `á, à, ả, ã, ạ`, etc.). Zero Windows API declarations (`user32.dll` removed), zero bitness branching, 100% portable.
* **Pipeline Sequence**:
  $$\text{Trim} \rightarrow \text{Collapse Whitespace} \rightarrow \text{Remove Whitespace before Punctuation} \rightarrow \text{Apply Casing} \rightarrow \text{Append Trailing Period}$$
* **Public Signature**:
  ```vb
  Public Function FormatText( _
      ByVal Text As String, _
      Optional ByVal Casing As VnCasingStyle = VnCaseSentence, _
      Optional ByVal AddPeriod As Boolean = True _
  ) As String
  ```

### 4.5 `CoreCoordinator.bas`
Unifies the transformation pipeline for callers.
* **Advanced & Simple Signatures**:
  ```vb
  Public Function ConvertNumber( _
      ByVal Value As Variant, _
      ByRef EngineOpts As VnEngineOptions, _
      ByRef CurrOpts As VnCurrencyOptions, _
      ByRef FormatOpts As VnFormatOptions, _
      ByRef OutErrorMessage As String _
  ) As String

  Public Function TryConvertNumber( _
      ByVal Value As Variant, _
      ByRef EngineOpts As VnEngineOptions, _
      ByRef CurrOpts As VnCurrencyOptions, _
      ByRef FormatOpts As VnFormatOptions, _
      ByRef OutText As String, _
      ByRef OutErrorMessage As String _
  ) As Boolean

  Public Function ConvertNumberDefault(ByVal Value As Variant) As String
  ```

---

## 5. Excel Interaction Layer (`src/excel/`)

### 5.1 `CellProcessor.bas`
Coordinates reading from worksheets, validating ranges, executing bulk transformations, and writing outputs.

#### A. Application State Preservation & Cleanup Contract
Caches initial states and restores them explicitly in the `Cleanup` block on both success and error paths:
```vb
Dim prevScreenUpdating As Boolean
Dim prevCalculation As XlCalculation
Dim prevEnableEvents As Boolean
Dim prevDisplayAlerts As Boolean
Dim prevDisplayStatusBar As Boolean
Dim prevStatusBarValue As Variant

On Error GoTo ErrorHandler

prevScreenUpdating = Application.ScreenUpdating
prevCalculation = Application.Calculation
prevEnableEvents = Application.EnableEvents
prevDisplayAlerts = Application.DisplayAlerts
prevDisplayStatusBar = Application.DisplayStatusBar
prevStatusBarValue = Application.StatusBar

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

' ... Processing ...

Cleanup:
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    Application.DisplayAlerts = prevDisplayAlerts
    Application.DisplayStatusBar = prevDisplayStatusBar
    Application.StatusBar = prevStatusBarValue
    On Error GoTo 0
```

#### B. Range Validation Rules
1. **Contiguous Single Area**: Rejects multi-area selections (`Range.Areas.Count <> 1`).
2. **Same Workbook**: Source and destination must belong to the same workbook (`SourceRange.Parent.Parent Is DestinationRange.Parent.Parent`). Cross-sheet within the same workbook is supported.
3. **No Overlap**: Source and destination ranges must not intersect (`Intersect(SourceRange, DestinationRange) Is Nothing`). Overlap is rejected for both Static and Formula modes.
4. **Destination Shape**:
   * Accepts a **single anchor cell** (e.g. `B2`), automatically expanding to match source dimensions.
   * Or accepts an **exact-size range** matching source rows and columns.
   * Partial/mismatched ranges are rejected.
5. **Worksheet Boundaries**: Validates that destination expansion does not exceed `Worksheet.Rows.Count` (1,048,576) or `Worksheet.Columns.Count` (16,384).
6. **Protected Sheet Check**: If destination worksheet is protected, verifies that all destination cells are unlocked. If any destination cell is locked, aborts atomically.
7. **Merged Cells**: Rejects merged cells in batch operations. Single source to single merged destination writes to `MergeArea.Cells(1, 1)`.
8. **Occupied Cell Detection**: Any cell containing a constant OR formula (even if returning `""`) is treated as occupied.

#### C. In-Memory Bulk Array Execution & Skip Preservation
* Reads source via `srcValues = SourceRange.Value2` (normalized to a 2D Variant array).
* Processes conversion in memory.
* **Skip Semantics**: Any cell skipped (empty, text error, invalid) **must never overwrite** the corresponding destination cell. The destination retains its original content.
* Performance benchmark targets: 1,000 cells $< 1$s; 10,000 cells $< 5$s.

#### D. Structured Batch Result
```vb
Public Type VnBatchResult
    Success As Boolean
    ConvertedCount As Long
    SkippedCount As Long
    ErrorCount As Long
    UndoAvailable As Boolean
    ErrorMessage As String
End Type
```

### 5.2 `UndoManager.bas`
Provides transactional 1-level Undo for conversions up to 10,000 cells.
* **Threshold Guard**:
  ```vb
  Public Const MAX_UNDO_CELLS As Long = 10000
  ```
  If cells exceed 10,000, prompts the user: conversion will proceed with Undo disabled.
* **Two-Phase Commit**:
  1. `PrepareUndoSnapshot(TargetRange)`: Captures `Workbook`, `Worksheet`, `Address`, `Values`, `Formulas`, and `HasFormula` flags into a staged buffer.
  2. `CommitUndoSnapshot()`: Only called after `ConvertRange` succeeds, replacing previous Undo state. If conversion fails, staged buffer is discarded and previous Undo remains intact.
* **Restoration**: Restores `.Formula` for formulas, `.Value2` for constants, and `ClearContents` for previously blank cells.
* **UI Invalidation**: Calls `gRibbon.InvalidateControl "btnUndo"` on all state transitions.

### 5.3 `Settings.bas`
* **Storage Location**: `HKCU\Software\VB and VBA Program Settings\BWPConvertTTNVN`.
* **Schema Versioning**: Includes `SettingsVersion = 1` for safe migration.
* **Contract**:
  * `LoadAllSettings`: Validates bounds of all enums; falls back to defaults if registry data is corrupt.
  * `SaveAllSettings`: Persists validated settings.
  * `ResetSettingsToDefault`: Clears registry and writes factory defaults.

### 5.4 `UDF.bas` — Worksheet Functions
Worksheet formulas are deterministic and **never read Registry settings**.
* **`=BWPVNWORDS(Target, [ZeroStyle], [ThousandStyle])`**: General number reading (Mode 2 decimals enabled by default).
* **`=BWPVND(Target, [AddChan], [ZeroStyle], [ThousandStyle])`**: Vietnamese Dong (truncate toward zero, adds `"đồng"`, optional `"chẵn"`).
* **`=BWPVNDUPPER(Target, [AddChan])`**: All-uppercase Vietnamese Dong.
* **Rules**:
  * Return type is `Variant`. Returns `CVErr(xlErrValue)` on invalid input or $> \text{MAX\_SUPPORTED\_VALUE}$.
  * Blank input returns `""` (empty string) to keep templates clean.
  * Dynamic arrays rejected in v1.0 (expects scalar input).
  * `Application.Volatile` is **omitted** (recalculates only when precedent cells change).

### 5.5 `RibbonCallbacks.bas`
* Caches `gRibbon As IRibbonUI` during `OnRibbonLoad`.
* Validates `Selection` is a single contiguous `Range` before dispatching.
* **Quick Convert Routing**:
  * 1D (cell, column, row) $\rightarrow$ executes Quick Convert.
  * 2D selection $\rightarrow$ redirects to `frmConvert` with source prefilled and destination left blank.
* Invalidation: Drives `GetUndoEnabled` callback dynamically.

---

## 6. UI Workflow & UserForm Specifications

### 6.1 Ribbon Definition (`ribbon/customUI14.xml`)
```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="tabBWPConvert" label="BWPConvertTTNVN" keytip="B">
        <group id="grpConversion" label="Chuyển đổi số">
          <button id="btnConvert" label="Đổi số thành chữ" size="large"
                  imageMso="ChangeTextCase" onAction="OnConvertClick"
                  supertip="Mở hộp thoại chuyển đổi số thành chữ tiếng Việt với đầy đủ tùy chọn." />
          <button id="btnQuickConvert" label="Chuyển nhanh" size="large"
                  imageMso="AutoSum" onAction="OnQuickConvertClick"
                  supertip="Tự động chuyển đổi số tại vùng chọn sang cột liền kề hoặc dòng dưới." />
          <button id="btnUndo" label="Hoàn tác" size="normal"
                  imageMso="Undo" onAction="OnUndoClick" getEnabled="GetUndoEnabled"
                  supertip="Khôi phục lại dữ liệu trước lần chuyển đổi gần nhất." />
        </group>
        <group id="grpSettings" label="Hệ thống">
          <button id="btnSettings" label="Cài đặt" size="normal"
                  imageMso="ControlProperties" onAction="OnSettingsClick"
                  supertip="Tùy chỉnh quy tắc đọc số, hướng chuyển nhanh và định dạng." />
          <button id="btnAbout" label="Giới thiệu" size="normal"
                  imageMso="Info" onAction="OnAboutClick"
                  supertip="Thông tin phiên bản, bản quyền và giấy phép mã nguồn mở." />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

### 6.2 Main Conversion Dialog (`frmConvert`)
* **State Retention**: Holds internal `mSourceRange As Excel.Range` and `mDestinationRange As Excel.Range` references (preserves worksheet identity across cross-sheet selections in the same workbook).
* **Range Picker Flow**:
  * Hides form $\rightarrow$ calls `Application.InputBox(..., Type:=8)`.
  * Catches cancellation cleanly without object assignment errors.
  * Verifies `pickedRange.Parent.Parent Is mHostWorkbook`.
  * Restores form.
* **Controls**:
  * Mutually exclusive Casing ComboBox: `cboCasing` (*Viết hoa chữ đầu*, *VIẾT HOA TOÀN BỘ*, *viết thường toàn bộ*).
  * Dynamic Currency Controls: Changing `cboCurrency` updates visible/enabled toggles for `"đồng"`, `"chẵn"`, sub-units.
  * Output Mode: `optStatic` (default) vs `optFormula` (generates namespaced `=BWPVND(...)`).
* **Copyright Notice**: Small footer label `Copyright © 2026 - IT Leon`.
* **Keyboard**: `Enter` $\rightarrow$ Convert; `Esc` / `X` $\rightarrow$ Cancel.

### 6.3 Settings Dialog (`frmSettings`)
* **Contract**:
  * `[ Mặc định ]`: Populates UI controls with factory defaults; **does not** touch the Registry.
  * `[ Lưu ]`: Validates and commits current UI state to Registry.
  * `[ Hủy ]` / `Esc` / `X`: Closes dialog with zero persistence changes.
* **Direction Enum**: `VnQuickAuto = 0` (column $\rightarrow$ right, row $\rightarrow$ below, cell $\rightarrow$ right), `VnQuickRight = 1`, `VnQuickBelow = 2`.

### 6.4 About Dialog (`frmAbout`)
* Displays version `1.0.0`, MIT License, and verifiable offline/privacy statements.
* **Mandatory Prominent UI Requirement**: Displays bold notice:
  **`Copyright © 2026 - IT Leon`**

---

## 7. Build, Test & Packaging Pipeline

### 7.1 Build Environment Contract
* **Rule**: *“Automated `.xlam` packaging requires Microsoft Excel for Windows on the build machine.”*
* **AccessVBOM Verification**:
  1. Registry check: `HKCU\Software\Microsoft\Office\<version>\Excel\Security\AccessVBOM`.
  2. Functional COM test: attempts accessing `Workbook.VBProject.VBComponents.Count`.
  3. Clear failure message with guidance if blocked. Optional build switch `-EnableAccessVBOM` temporarily enables the flag and restores original value in a `finally` block.

### 7.2 Isolated Process Execution & PID Tracking
* Uses `win32com.client.DispatchEx("Excel.Application")` (isolated hidden instance).
* Obtains Excel Windows handle `excel.Hwnd` and resolves its exact Windows Process ID (`PID`).
* **Process Safety**: Cleanup code terminates **only** the build-owned PID. Never executes global `Stop-Process -Name EXCEL`.

### 7.3 Multi-Stage Build Pipeline (`scripts/build.py`)
1. **Stage 0 — Version & Contract Check**: Reads `VERSION`, validates `.frm` + `.frx` pairs, verifies Ribbon callback signatures match `RibbonCallbacks.bas`.
2. **Stage 1 — Environment Preflight**: Checks Windows, Excel installation, AccessVBOM capability.
3. **Stage 2 — Isolated COM Build**: Headless Excel creates workbook, sets Add-in metadata, imports all `.bas` and `.frm` components, runs smoke compile validation, saves as `dist/BWPConvertTTNVN.xlam` (`FileFormat = 55`). Cleanly terminates COM process.
4. **Stage 3 — Direct OpenXML Ribbon Injection (`scripts/package_ribbon.py`)**:
   * Opens `.xlam` as ZIP archive.
   * Injects `customUI/customUI14.xml`.
   * Adds relationship `<Relationship Id="rIdCustomUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>` into `_rels/.rels`.
   * Verifies `[Content_Types].xml` and `xl/vbaProject.bin` integrity. Re-zips cleanly.
5. **Stage 4 — Reopen Validation**: Fresh isolated Excel COM instance opens the final packaged `.xlam` to verify OpenXML structural validity.
6. **Stage 5 — Automated Test Execution (`tests/run_tests.py`)**:
   * **Test Class A (Core API)**: Calls `'BWPConvertTTNVN.xlam'!ConvertNumberDefault` via `excel.Run`.
   * **Test Class B (Worksheet UDF)**: Writes actual formula `=BWPVND(A1)` to a temporary `.xlsx`, calls `excel.CalculateFull()`, and asserts cell value.
   * **Test Class C (Integration)**: Tests Quick Convert, batch arrays, destination expansion, overlap rejection, Undo.
   * Driven by `tests/expected_cases.csv`.
7. **Stage 6 — Release Packaging**:
   * Bundles `BWPConvertTTNVN.xlam`, `Install.vbs`, `README.md`, `LICENSE`, `CHANGELOG.md` into `dist/BWPConvertTTNVN-v1.0.0.zip`.
   * Generates cryptographic checksum: `dist/BWPConvertTTNVN-v1.0.0.sha256`.

### 7.4 Installer Contract (`installer/Install.vbs`)
* Non-elevated installer running under standard user permissions.
* **Process Detection**: Checks if Excel is currently running. If open, alerts user to save work and closes gracefully.
* **Mark of the Web (MOTW) Removal**: Microsoft blocks downloaded Internet macros by default. The installer explicitly unblocks the add-in file by deleting the `Zone.Identifier` alternate data stream before registration.
* **Idempotent Registration**: Inspects `Application.AddIns`. Reuses existing registration if present, otherwise calls `AddIns.Add(destPath, True)`, setting `Installed = True`.

---

## 8. Acceptance Criteria & Test Matrix

| ID | Category | Test Condition / Input | Expected Result |
|:---|:---|:---|:---|
| **AC-01** | Zero | `0` | `"Không đồng chẵn."` (VND) / `"không"` (Pure) |
| **AC-02** | Basic | `1`, `5`, `10`, `11`, `15`, `20`, `21`, `24`, `25` | All irregular rules (*mười một*, *mười lăm*, *hai mươi mốt*, *hai mươi tư*, *hai mươi lăm*) pass |
| **AC-03** | Hundred | `105`, `110`, `115`, `121`, `125` | *một trăm lẻ năm*, *một trăm mười*, *một trăm hai mươi mốt* |
| **AC-04** | Scale 10^3 | `1,000`, `1,001`, `1,005`, `1,010`, `1,100` | *một nghìn không trăm lẻ một*, *một nghìn không trăm lẻ năm* |
| **AC-05** | Scale 10^6 | `1,000,000`, `1,000,001`, `1,000,005`, `1,000,010`, `1,000,100`, `1,001,000`, `1,050,000`, `125,430,000` | Triplet 000 skipped; `forceFullTriplet` produces *không trăm lẻ năm* |
| **AC-06** | Scale 10^9 | `1,000,000,000`, `1,000,000,001` | *một tỷ không trăm lẻ một* |
| **AC-07** | Scale 10^12 | `1,000,000,000,000`, `1,000,000,000,001` | *một nghìn tỷ không trăm lẻ một* |
| **AC-08** | Boundaries | `999,999,999,999,998`, `999,999,999,999,999` (`MAX_SUPPORTED_VALUE`), `-999,999,999,999,999` | Successfully converted |
| **AC-09** | Overflow | `1,000,000,000,000,000` | Returns `CVErr(xlErrValue)` in UDF / validation error in GUI |
| **AC-10** | Negative | `-150,000` | `"Âm một trăm năm mươi nghìn đồng chẵn."` |
| **AC-11** | Decimals | `125.5` (VND) $\rightarrow$ *Một trăm hai mươi lăm đồng chẵn.*; `125.05` (General) $\rightarrow$ *Một trăm hai mươi lăm phẩy không năm.* | Mode 1 truncates toward zero; Mode 2 spells decimal digits |
| **AC-12** | Dialects | `ZeroStyle = linh`, `ThousandStyle = ngàn` | Outputs *linh* and *ngàn* respectively |
| **AC-13** | Batch | Range `A2:A100` with 1 anchor `B2` | Converted in $< 1$s; skipped cells leave destination untouched |
| **AC-14** | Overlap | Source `A1:A10`, Dest `A5:A14` | Rejected with clear overlap warning |
| **AC-15** | Undo | Convert 500 cells $\rightarrow$ click Undo | 100% original values and formulas restored |
| **AC-16** | Worksheet UDF | `=BWPVND(A1)` in live workbook | Calculates properly without depending on Registry settings |
| **AC-17** | Copyright UI | Inspect `frmAbout` and `frmConvert` | `Copyright © 2026 - IT Leon` is clearly and visibly present |
| **AC-18** | Installer | Execute `Install.vbs` on fresh system | Unblocks MOTW, registers add-in, Ribbon tab appears on Excel launch |
