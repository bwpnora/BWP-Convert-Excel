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
        MsgBox "Please select a cell or data range.", vbExclamation, BuildInfo.APP_NAME
        Exit Sub
    End If
    On Error GoTo 0
    frmConvert.Show
End Sub

Public Sub OnQuickConvertClick(ByVal control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell or data range.", vbExclamation, BuildInfo.APP_NAME
        Exit Sub
    End If
    
    Dim srcRng As Range
    Set srcRng = Selection
    If srcRng Is Nothing Then Exit Sub
    
    If srcRng.Areas.Count <> 1 Then
        MsgBox "Multi-area selections are not supported.", vbExclamation, BuildInfo.APP_NAME
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
    
    ' Quick Convert executes in-place replacement on the selected range
    Dim engOpts As VnEngineOptions
    Dim currOpts As VnCurrencyOptions
    Dim fmtOpts As VnFormatOptions
    Dim appOpts As VnAppSettings
    Settings.LoadAllSettings engOpts, currOpts, fmtOpts, appOpts
    
    ' In-place quick conversion requires static text mode (avoiding circular formula references)
    appOpts.OutputMode = VnOutputStatic
    
    Dim destRng As Range
    Set destRng = srcRng
    
    Dim batchResult As VnBatchResult
    batchResult = CellProcessor.ConvertRange(srcRng, destRng, engOpts, currOpts, fmtOpts, appOpts)
    
    If Not batchResult.Success Then
        MsgBox batchResult.ErrorMessage, vbExclamation, BuildInfo.APP_NAME
        Exit Sub
    End If
    
    InvalidateRibbonControl "btnUndo"
    
    If appOpts.ShowBatchSummary And batchResult.ConvertedCount > 1 Then
        MsgBox "Successfully converted " & batchResult.ConvertedCount & " cell(s)." & _
               IIf(batchResult.SkippedCount > 0, vbCrLf & "Skipped: " & batchResult.SkippedCount & " cell(s).", ""), _
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
