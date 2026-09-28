Attribute VB_Name = "CellProcessor"
Option Explicit

' ==============================================================================
' CellProcessor.bas
' Excel batch range processor for BWPConvertTTNVN.
' Features:
' 1. Robust state caching and guaranteed restoration on success and error paths.
' 2. Range validation: single area, same workbook, cross-sheet safe overlap.
' 3. Destination shape handling: single-cell anchor auto-expansion, exact matching.
' 4. Safety checks: worksheet boundaries, merged cells guard, locked cells guard.
' 5. Formula-safe skip preservation: skipped cells never overwrite destination.
' 6. Transactional 2-phase Undo coordination with UndoManager.
' Pure 7-bit ASCII safe source code.
' ==============================================================================

' Structured Batch Result UDT
Public Type VnBatchResult
    Success As Boolean
    ConvertedCount As Long
    SkippedCount As Long
    ErrorCount As Long
    UndoAvailable As Boolean
    ErrorMessage As String
End Type

' ==============================================================================
' Core Public Entry Point: ConvertRange
' ==============================================================================

Public Function ConvertRange( _
    ByVal SourceRange As Range, _
    ByVal DestinationRange As Range, _
    ByRef EngineOpts As VnEngineOptions, _
    ByRef CurrOpts As VnCurrencyOptions, _
    ByRef FormatOpts As VnFormatOptions, _
    ByRef AppOpts As VnAppSettings _
) As VnBatchResult
    Dim result As VnBatchResult
    result.Success = False
    result.ConvertedCount = 0
    result.SkippedCount = 0
    result.ErrorCount = 0
    result.UndoAvailable = False
    result.ErrorMessage = ""

    ' State caching variables
    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation
    Dim prevEnableEvents As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevDisplayStatusBar As Boolean
    Dim prevStatusBarValue As Variant
    Dim stateCached As Boolean
    stateCached = False

    On Error GoTo ErrorHandler

    ' 1. Cache application state
    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation
    prevEnableEvents = Application.EnableEvents
    prevDisplayAlerts = Application.DisplayAlerts
    prevDisplayStatusBar = Application.DisplayStatusBar
    prevStatusBarValue = Application.StatusBar
    stateCached = True

    ' 2. Validate range objects exist
    If SourceRange Is Nothing Or DestinationRange Is Nothing Then
        result.ErrorMessage = "Vung du lieu nguon hoac dich khong hop le."
        GoTo Cleanup
    End If

    ' 3. Validate single contiguous area (no multi-area selections)
    If SourceRange.Areas.Count <> 1 Or DestinationRange.Areas.Count <> 1 Then
        result.ErrorMessage = "Khong ho tro vung chon khong lien tuc (multi-area)."
        GoTo Cleanup
    End If

    ' 4. Validate same workbook
    If Not (SourceRange.Parent.Parent Is DestinationRange.Parent.Parent) Then
        result.ErrorMessage = "Vung nguon va vung dich phai thuoc cung mot workbook."
        GoTo Cleanup
    End If

    ' 5. Destination dimension expansion and validation
    Dim srcRows As Long, srcCols As Long
    srcRows = SourceRange.Rows.Count
    srcCols = SourceRange.Columns.Count

    Dim destRows As Long, destCols As Long
    destRows = DestinationRange.Rows.Count
    destCols = DestinationRange.Columns.Count

    Dim destWs As Worksheet
    Set destWs = DestinationRange.Parent

    Dim destExpanded As Range
    Dim destStartRow As Long, destStartCol As Long
    destStartRow = DestinationRange.Row
    destStartCol = DestinationRange.Column

    If destRows = 1 And destCols = 1 Then
        ' Single anchor cell -> automatically expands to match source dimensions
        If (destStartRow + srcRows - 1 > destWs.Rows.Count) Or _
           (destStartCol + srcCols - 1 > destWs.Columns.Count) Then
            result.ErrorMessage = "Vung dich vuot qua gioi han kich thuoc cua worksheet."
            GoTo Cleanup
        End If
        Set destExpanded = destWs.Range( _
            destWs.Cells(destStartRow, destStartCol), _
            destWs.Cells(destStartRow + srcRows - 1, destStartCol + srcCols - 1) _
        )
    Else
        ' Multi-cell destination must match source rows and columns exactly
        If destRows <> srcRows Or destCols <> srcCols Then
            result.ErrorMessage = "Kich thuoc vung dich khong khop voi vung nguon."
            GoTo Cleanup
        End If
        Set destExpanded = DestinationRange
    End If

    ' 6. Cross-sheet safe overlap check
    Dim sameSheet As Boolean
    sameSheet = (SourceRange.Parent Is destWs)

    If sameSheet Then
        Dim isect As Range
        Set isect = Nothing
        Set isect = Application.Intersect(SourceRange, destExpanded)
        If Not (isect Is Nothing) Then
            result.ErrorMessage = "Vung nguon va vung dich bi trung hoac de len nhau tren cung sheet."
            GoTo Cleanup
        End If
    End If

    ' 7. Merged cells check: reject if batch operation (> 1 cell) has merged cells
    If srcRows * srcCols > 1 Then
        Dim srcMerged As Boolean
        Dim destMerged As Boolean
        
        srcMerged = False
        If IsNull(SourceRange.MergeCells) Then
            srcMerged = True
        ElseIf CBool(SourceRange.MergeCells) Then
            srcMerged = True
        End If
        
        destMerged = False
        If IsNull(destExpanded.MergeCells) Then
            destMerged = True
        ElseIf CBool(destExpanded.MergeCells) Then
            destMerged = True
        End If
        
        If srcMerged Or destMerged Then
            result.ErrorMessage = "Khong ho tro chuyen doi hang loat tren vung co o bi tron (merged cells)."
            GoTo Cleanup
        End If
    End If

    ' 8. Protected sheet check: abort atomically if any destination cell is locked
    If destWs.ProtectContents Then
        Dim destLocked As Variant
        destLocked = destExpanded.Locked
        If IsNull(destLocked) Then
            result.ErrorMessage = "Worksheet dich bi khoa va co chua o bi khoa."
            GoTo Cleanup
        ElseIf CBool(destLocked) Then
            result.ErrorMessage = "Worksheet dich bi khoa va vung dich bi khoa."
            GoTo Cleanup
        End If
    End If

    ' 9. Formula mode validation: reject USD and Custom currency
    Dim isFormulaMode As Boolean
    isFormulaMode = (AppOpts.OutputMode = VnOutputFormula)

    If isFormulaMode Then
        If CurrOpts.CurrencyType = VnCurrUSD Or CurrOpts.CurrencyType = VnCurrCustom Then
            result.ErrorMessage = "Che do cong thuc khong ho tro loai tien te USD hoac tuy chinh."
            GoTo Cleanup
        End If
    End If

    ' 10. Suppress screen updating and events for processing
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' 11. Prepare Undo snapshot
    UndoManager.PrepareUndoSnapshot destExpanded

    ' 12. Read source values into normalized 2D array
    Dim srcValues As Variant
    If srcRows = 1 And srcCols = 1 Then
        Dim singleArr(1 To 1, 1 To 1) As Variant
        singleArr(1, 1) = SourceRange.Value2
        srcValues = singleArr
    Else
        srcValues = SourceRange.Value2
    End If

    ' 13. Process batch conversion with formula-safe skip preservation
    Dim r As Long, c As Long
    Dim convertedCount As Long
    Dim skippedCount As Long
    Dim errorCount As Long

    convertedCount = 0
    skippedCount = 0
    errorCount = 0

    For r = 1 To srcRows
        For c = 1 To srcCols
            Dim v As Variant
            v = srcValues(r, c)

            Dim isSkipped As Boolean
            isSkipped = False

            If IsEmpty(v) Or IsNull(v) Or IsError(v) Then
                isSkipped = True
            ElseIf VarType(v) = vbBoolean Or VarType(v) = vbDate Then
                isSkipped = True
            ElseIf Not IsNumeric(v) Then
                isSkipped = True
            End If

            If isSkipped Then
                skippedCount = skippedCount + 1
                ' Formula-safe skip preservation:
                ' Destination cell is never touched. Existing values and formulas remain intact.
            Else
                Dim targetCell As Range
                If srcRows = 1 And srcCols = 1 And destExpanded.MergeCells Then
                    Set targetCell = destExpanded.MergeArea.Cells(1, 1)
                Else
                    Set targetCell = destExpanded.Cells(r, c)
                End If

                If isFormulaMode Then
                    Dim srcCell As Range
                    Set srcCell = SourceRange.Cells(r, c)
                    Dim cellRef As String

                    If sameSheet Then
                        cellRef = srcCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    Else
                        cellRef = "'" & Replace(SourceRange.Parent.Name, "'", "''") & "'!" & _
                                  srcCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    End If

                    Dim fmlaStr As String
                    If CurrOpts.CurrencyType = VnCurrVND Then
                        If FormatOpts.Casing = VnCaseUpper Then
                            fmlaStr = "=BWPVNDUPPER(" & cellRef & ", " & IIf(CurrOpts.AddChan, "TRUE", "FALSE") & ")"
                        Else
                            fmlaStr = "=BWPVND(" & cellRef & ", " & IIf(CurrOpts.AddChan, "TRUE", "FALSE") & ", " & _
                                      CStr(CLng(EngineOpts.ZeroStyle)) & ", " & CStr(CLng(EngineOpts.ThousandStyle)) & ")"
                        End If
                    Else ' VnCurrNone
                        fmlaStr = "=BWPVNWORDS(" & cellRef & ", " & CStr(CLng(EngineOpts.ZeroStyle)) & ", " & _
                                  CStr(CLng(EngineOpts.ThousandStyle)) & ")"
                    End If

                    targetCell.Formula = fmlaStr
                    convertedCount = convertedCount + 1
                Else
                    ' Static Mode
                    Dim outText As String
                    Dim outErr As String
                    If TryConvertNumber(v, EngineOpts, CurrOpts, FormatOpts, outText, outErr) Then
                        targetCell.Value2 = outText
                        convertedCount = convertedCount + 1
                    Else
                        errorCount = errorCount + 1
                    End If
                End If
            End If
        Next c
    Next r

    ' 14. Commit Undo snapshot on success
    If convertedCount > 0 Then
        UndoManager.CommitUndoSnapshot
    Else
        UndoManager.DiscardStagedSnapshot
    End If

    result.Success = True
    result.ConvertedCount = convertedCount
    result.SkippedCount = skippedCount
    result.ErrorCount = errorCount
    result.UndoAvailable = UndoManager.IsUndoAvailable()
    result.ErrorMessage = ""

Cleanup:
    On Error Resume Next
    If stateCached Then
        Application.ScreenUpdating = prevScreenUpdating
        Application.Calculation = prevCalculation
        Application.EnableEvents = prevEnableEvents
        Application.DisplayAlerts = prevDisplayAlerts
        Application.DisplayStatusBar = prevDisplayStatusBar
        Application.StatusBar = prevStatusBarValue
    End If
    On Error GoTo 0
    ConvertRange = result
    Exit Function

ErrorHandler:
    result.Success = False
    result.ErrorMessage = Err.Description
    If Len(result.ErrorMessage) = 0 Then
        result.ErrorMessage = "Loi xu ly CellProcessor: " & CStr(Err.Number)
    End If
    UndoManager.DiscardStagedSnapshot
    Resume Cleanup
End Function

' ==============================================================================
' Occupied Cell Detection Helper
' ==============================================================================

Public Function HasOccupiedCells(ByVal TargetRange As Range) As Boolean
    If TargetRange Is Nothing Then
        HasOccupiedCells = False
        Exit Function
    End If

    If TargetRange.Rows.Count = 1 And TargetRange.Columns.Count = 1 Then
        HasOccupiedCells = (Not IsEmpty(TargetRange.Value2)) Or TargetRange.HasFormula
        Exit Function
    End If

    Dim hasConst As Boolean
    Dim hasFmla As Boolean

    On Error Resume Next
    hasConst = Not (TargetRange.SpecialCells(xlCellTypeConstants) Is Nothing)
    hasFmla = Not (TargetRange.SpecialCells(xlCellTypeFormulas) Is Nothing)
    On Error GoTo 0

    HasOccupiedCells = hasConst Or hasFmla
End Function

' ==============================================================================
' Scalar Bridge Overloads for COM Automation Testing
' ==============================================================================

Public Function ConvertRangeBridge( _
    ByVal SourceRange As Range, _
    ByVal DestinationRange As Range, _
    Optional ByVal OutputMode As Long = 0, _
    Optional ByVal CurrType As Long = 0, _
    Optional ByVal AddChan As Boolean = True, _
    Optional ByVal DecPlaces As Long = 0, _
    Optional ByVal ZeroStyle As Long = 0, _
    Optional ByVal ThousandStyle As Long = 0, _
    Optional ByVal FourStyle As Long = 0, _
    Optional ByVal DecMode As Long = 0, _
    Optional ByVal Casing As Long = 0, _
    Optional ByVal AddPeriod As Boolean = True _
) As Variant
    Dim engOpts As VnEngineOptions
    Dim currOpts As VnCurrencyOptions
    Dim fmtOpts As VnFormatOptions
    Dim appOpts As VnAppSettings

    engOpts.ZeroStyle = ZeroStyle
    engOpts.ThousandStyle = ThousandStyle
    engOpts.FourStyle = FourStyle
    engOpts.DecimalMode = DecMode

    currOpts.CurrencyType = CurrType
    currOpts.AddChan = AddChan
    currOpts.DecimalPlaces = CInt(DecPlaces)
    currOpts.CustomPrefix = ""
    currOpts.CustomSuffix = ""
    currOpts.CustomSubUnit = ""

    fmtOpts.Casing = Casing
    fmtOpts.AddPeriod = AddPeriod

    appOpts.OutputMode = OutputMode
    appOpts.QuickConvertDirection = VnQuickAuto
    appOpts.ConfirmOverwrite = True
    appOpts.ShowBatchSummary = True
    appOpts.SettingsVersion = 1

    Dim res As VnBatchResult
    res = ConvertRange(SourceRange, DestinationRange, engOpts, currOpts, fmtOpts, appOpts)

    ConvertRangeBridge = Array( _
        res.Success, _
        res.ConvertedCount, _
        res.SkippedCount, _
        res.ErrorCount, _
        res.UndoAvailable, _
        res.ErrorMessage _
    )
End Function

Public Function ConvertRangeWithSavedSettings( _
    ByVal SourceRange As Range, _
    ByVal DestinationRange As Range _
) As Variant
    Dim engOpts As VnEngineOptions
    Dim currOpts As VnCurrencyOptions
    Dim fmtOpts As VnFormatOptions
    Dim appOpts As VnAppSettings

    LoadAllSettings engOpts, currOpts, fmtOpts, appOpts

    Dim res As VnBatchResult
    res = ConvertRange(SourceRange, DestinationRange, engOpts, currOpts, fmtOpts, appOpts)

    ConvertRangeWithSavedSettings = Array( _
        res.Success, _
        res.ConvertedCount, _
        res.SkippedCount, _
        res.ErrorCount, _
        res.UndoAvailable, _
        res.ErrorMessage _
    )
End Function
