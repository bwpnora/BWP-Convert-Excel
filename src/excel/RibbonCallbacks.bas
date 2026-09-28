Attribute VB_Name = "RibbonCallbacks"
Option Explicit

' ==============================================================================
' RibbonCallbacks.bas
' CustomUI Ribbon callbacks and actions for BWPConvertTTNVN.
' Office 2010+ CustomUI namespace: http://schemas.microsoft.com/office/2009/07/customui
' Pure 7-bit ASCII safe source code.
' ==============================================================================

Private gRibbon As Object

Public Property Get RibbonHandle() As Object
    Set RibbonHandle = gRibbon
End Property

Public Property Set RibbonHandle(ByVal ribbon As Object)
    Set gRibbon = ribbon
End Property

' ==============================================================================
' Ribbon Lifecycle
' ==============================================================================

Public Sub OnRibbonLoad(ByVal ribbon As Object)
    Set gRibbon = ribbon
End Sub

Public Sub InvalidateRibbonControl(ByVal controlId As String)
    On Error Resume Next
    If Not gRibbon Is Nothing Then
        gRibbon.InvalidateControl controlId
    End If
    On Error GoTo 0
End Sub

Public Sub InvalidateUndoButton()
    InvalidateRibbonControl "btnUndo"
End Sub

Public Sub InvalidateUndoRibbon()
    InvalidateRibbonControl "btnUndo"
End Sub

' ==============================================================================
' Ribbon Button Actions
' ==============================================================================

Public Sub OnConvertClick(ByVal control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then
        MsgBox "Vui l" & ChrW$(&HF2) & "ng ch" & ChrW$(&H1ECD) & "n m" & ChrW$(&H1ED9) & "t " & ChrW$(&HF4) & " ho" & ChrW$(&H1EA1) & "c v" & ChrW$(&HF9) & "ng d" & ChrW$(&H1EEF) & " li" & ChrW$(&H1EC7) & "u.", vbExclamation, BuildInfo.APP_NAME
        Exit Sub
    End If
    On Error GoTo 0
    frmConvert.Show
End Sub

Public Sub OnQuickConvertClick(ByVal control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then
        MsgBox "Vui l" & ChrW$(&HF2) & "ng ch" & ChrW$(&H1ECD) & "n m" & ChrW$(&H1ED9) & "t " & ChrW$(&HF4) & " ho" & ChrW$(&H1EA1) & "c v" & ChrW$(&HF9) & "ng d" & ChrW$(&H1EEF) & " li" & ChrW$(&H1EC7) & "u.", vbExclamation, BuildInfo.APP_NAME
        Exit Sub
    End If
    
    Dim srcRng As Range
    Set srcRng = Selection
    If srcRng Is Nothing Then Exit Sub
    
    If srcRng.Areas.Count <> 1 Then
        MsgBox "Kh" & ChrW$(&HF4) & "ng h" & ChrW$(&H1ED7) & " tr" & ChrW$(&H1EE3) & " v" & ChrW$(&HF9) & "ng ch" & ChrW$(&H1ECD) & "n kh" & ChrW$(&HF4) & "ng li" & ChrW$(&HEA) & "n t" & ChrW$(&H1EE5) & "c (multi-area).", vbExclamation, BuildInfo.APP_NAME
        Exit Sub
    End If
    On Error GoTo 0
    
    ' If 2D selection: open frmConvert with source prefilled and destination left blank
    If srcRng.Rows.Count > 1 And srcRng.Columns.Count > 1 Then
        Set frmConvert.SourceRange = srcRng
        Set frmConvert.DestinationRange = Nothing
        frmConvert.Show
        Exit Sub
    End If
    
    ' 1D selection: execute quick conversion
    Dim engOpts As VnEngineOptions
    Dim currOpts As VnCurrencyOptions
    Dim fmtOpts As VnFormatOptions
    Dim appOpts As VnAppSettings
    Settings.LoadAllSettings engOpts, currOpts, fmtOpts, appOpts
    
    Dim targetDir As VnQuickDirection
    targetDir = appOpts.QuickConvertDirection
    
    If targetDir = VnQuickAuto Then
        If srcRng.Rows.Count = 1 And srcRng.Columns.Count > 1 Then
            targetDir = VnQuickBelow
        Else
            targetDir = VnQuickRight
        End If
    End If
    
    Dim ws As Worksheet
    Set ws = srcRng.Parent
    
    Dim srcRows As Long, srcCols As Long
    srcRows = srcRng.Rows.Count
    srcCols = srcRng.Columns.Count
    
    Dim destRng As Range
    
    If targetDir = VnQuickRight Then
        If CLng(srcRng.Column) + srcCols + srcCols - 1 > ws.Columns.Count Then
            MsgBox "V" & ChrW$(&HF9) & "ng xu" & ChrW$(&H1EA5) & "t k" & ChrW$(&H1EBF) & "t qu" & ChrW$(&H1EA3) & " v" & ChrW$(&H1B0) & ChrW$(&H1EE3) & "t qu" & ChrW$(&HE1) & " gi" & ChrW$(&H1EDB) & "i h" & ChrW$(&H1EA1) & "n c" & ChrW$(&H1EE7) & "a b" & ChrW$(&H1EA3) & "ng t" & ChrW$(&HED) & "nh.", vbExclamation, BuildInfo.APP_NAME
            Exit Sub
        End If
        Set destRng = ws.Cells(srcRng.Row, srcRng.Column + srcCols).Resize(srcRows, srcCols)
    Else
        If CLng(srcRng.Row) + srcRows + srcRows - 1 > ws.Rows.Count Then
            MsgBox "V" & ChrW$(&HF9) & "ng xu" & ChrW$(&H1EA5) & "t k" & ChrW$(&H1EBF) & "t qu" & ChrW$(&H1EA3) & " v" & ChrW$(&H1B0) & ChrW$(&H1EE3) & "t qu" & ChrW$(&HE1) & " gi" & ChrW$(&H1EDB) & "i h" & ChrW$(&H1EA1) & "n c" & ChrW$(&H1EE7) & "a b" & ChrW$(&H1EA3) & "ng t" & ChrW$(&HED) & "nh.", vbExclamation, BuildInfo.APP_NAME
            Exit Sub
        End If
        Set destRng = ws.Cells(srcRng.Row + srcRows, srcRng.Column).Resize(srcRows, srcCols)
    End If
    
    If appOpts.ConfirmOverwrite Then
        If Application.WorksheetFunction.CountA(destRng) > 0 Then
            Dim ans As VbMsgBoxResult
            ans = MsgBox("V" & ChrW$(&HF9) & "ng " & ChrW$(&H111) & ChrW$(&HED) & "ch " & ChrW$(&H111) & ChrW$(&HE3) & " c" & ChrW$(&HF3) & " d" & ChrW$(&H1EEF) & " li" & ChrW$(&H1EC7) & "u. B" & ChrW$(&H1EA1) & "n c" & ChrW$(&HF3) & " ch" & ChrW$(&H1EAF) & "c ch" & ChrW$(&H1EAF) & "n mu" & ChrW$(&H1ED1) & "n ghi " & ChrW$(&H111) & ChrW$(&HE8) & " kh" & ChrW$(&HF4) & "ng?", vbQuestion + vbYesNo + vbDefaultButton2, BuildInfo.APP_NAME)
            If ans <> vbYes Then Exit Sub
        End If
    End If
    
    Dim batchResult As VnBatchResult
    batchResult = CellProcessor.ConvertRange(srcRng, destRng, engOpts, currOpts, fmtOpts, appOpts)
    
    If Not batchResult.Success Then
        MsgBox batchResult.ErrorMessage, vbExclamation, BuildInfo.APP_NAME
        Exit Sub
    End If
    
    InvalidateRibbonControl "btnUndo"
    
    If appOpts.ShowBatchSummary And batchResult.ConvertedCount > 1 Then
        MsgBox ChrW$(&H110) & ChrW$(&HE3) & " chuy" & ChrW$(&H1EC3) & "n " & ChrW$(&H111) & ChrW$(&H1ED5) & "i th" & ChrW$(&HE0) & "nh c" & ChrW$(&HF4) & "ng " & batchResult.ConvertedCount & " " & ChrW$(&HF4) & "." & _
               IIf(batchResult.SkippedCount > 0, vbCrLf & "B" & ChrW$(&H1ECF) & " qua: " & batchResult.SkippedCount & " " & ChrW$(&HF4) & ".", ""), _
               vbInformation, BuildInfo.APP_NAME
    End If
End Sub

Public Sub OnUndoClick(ByVal control As Object)
    UndoManager.ExecuteUndo
    InvalidateRibbonControl "btnUndo"
End Sub

Public Sub OnSettingsClick(ByVal control As Object)
    frmSettings.Show
End Sub

Public Sub OnAboutClick(ByVal control As Object)
    frmAbout.Show
End Sub

Public Sub GetUndoEnabled(ByVal control As Object, ByRef enabled As Variant)
    enabled = UndoManager.IsUndoAvailable()
End Sub
