VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConvert 
   Caption         =   "UserForm1"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7365
   OleObjectBlob   =   "frmConvert.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' frmConvert.frm
' Main conversion dialog for BWPConvertTTNVN.
' Pure 7-bit ASCII safe source code.
' Features:
' - Range object retention across cross-sheet selections in the same workbook.
' - InputBox(Type:=8) cancel handling without runtime errors.
' - Dynamic currency change: VND (formula/dong/chan enabled), USD (formula disabled), None.
' - Mandatory copyright notice: BuildInfo.APP_COPYRIGHT (Copyright (C) 2026 - IT Leon).
' - Pre-flight validations: contiguous area, same workbook, overlap, protected cells.
' ==============================================================================

Private mSourceRange As Excel.Range
Private mDestinationRange As Excel.Range
Private mHostWorkbook As Excel.Workbook
Private mBatchResult As VnBatchResult

Public Property Get SourceRange() As Excel.Range
    Set SourceRange = mSourceRange
End Property

Public Property Set SourceRange(ByVal rng As Excel.Range)
    Set mSourceRange = rng
    If Not rng Is Nothing Then
        Set mHostWorkbook = rng.Parent.Parent
        txtSourceRange.Text = GetQualifiedAddress(rng)
    Else
        txtSourceRange.Text = ""
    End If
End Property

Public Property Get DestinationRange() As Excel.Range
    Set DestinationRange = mDestinationRange
End Property

Public Property Set DestinationRange(ByVal rng As Excel.Range)
    Set mDestinationRange = rng
    If Not rng Is Nothing Then
        If mHostWorkbook Is Nothing Then
            Set mHostWorkbook = rng.Parent.Parent
        End If
        txtDestRange.Text = GetQualifiedAddress(rng)
    Else
        txtDestRange.Text = ""
    End If
End Property

Public Property Get LastBatchResult() As VnBatchResult
    LastBatchResult = mBatchResult
End Property

Public Property Get CurrencyIndex() As Long
    CurrencyIndex = cboCurrency.ListIndex
End Property

Public Property Let CurrencyIndex(ByVal val As Long)
    cboCurrency.ListIndex = val
    cboCurrency_Change
End Property

Public Property Get CasingIndex() As Long
    CasingIndex = cboCasing.ListIndex
End Property

Public Property Let CasingIndex(ByVal val As Long)
    cboCasing.ListIndex = val
End Property

Public Property Get FormulaModeEnabled() As Boolean
    FormulaModeEnabled = optFormula.Enabled
End Property

Public Property Get OutputModeValue() As Long
    If optFormula.Value Then OutputModeValue = 1 Else OutputModeValue = 0
End Property

Public Property Let OutputModeValue(ByVal val As Long)
    If val = 1 Then optFormula.Value = True Else optStatic.Value = True
End Property

Public Property Get TitleText() As String
    TitleText = ChrW$(&H0110) & ChrW$(&H1ED4) & "I S" & ChrW$(&H1ED0) & " TH" & ChrW$(&H00C0) & "NH CH" & ChrW$(&H1EEE) & " TI" & ChrW$(&H1EBE) & "NG VI" & ChrW$(&H1EC6) & "T"
End Property

Public Property Get CopyrightText() As String
    CopyrightText = lblCopyright.Caption
End Property

Public Property Get FormulaTipText() As String
    FormulaTipText = lblFormulaTip.Caption
End Property

Public Property Get AddDongValue() As Boolean
    AddDongValue = chkAddDong.Value
End Property

Public Property Get AddChanValue() As Boolean
    AddChanValue = chkAddChan.Value
End Property

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnConvert_Click()
    If ExecuteConversion(SuppressPrompts:=False) Then
        Unload Me
    End If
End Sub

Private Sub UserForm_Initialize()
    PopulateCaptions
    LoadDefaultState
End Sub

Public Sub PopulateCaptions()
    ' Title: DOI SO THANH CHU TIENG VIET
    Me.Caption = ChrW$(&H110) & ChrW$(&H1ED4) & "I S" & ChrW$(&H1ED0) & " TH" & ChrW$(&HC0) & "NH CH" & ChrW$(&H1EEE) & " TI" & ChrW$(&H1EBE) & "NG VI" & ChrW$(&H1EC6) & "T"
    
    lblSourceRange.Caption = "V" & ChrW$(&HF9) & "ng d" & ChrW$(&H1EEF) & " li" & ChrW$(&H1EC7) & "u ngu" & ChrW$(&H1ED3) & "n (ch" & ChrW$(&H1EE9) & "a s" & ChrW$(&H1ED1) & "):"
    btnSelectSource.Caption = "Ch" & ChrW$(&H1ECD) & "n..."
    
    lblDestRange.Caption = "V" & ChrW$(&HF9) & "ng xu" & ChrW$(&H1EA5) & "t k" & ChrW$(&H1EBF) & "t qu" & ChrW$(&H1EA3) & " (" & ChrW$(&HF4) & " b" & ChrW$(&H1EAF) & "t " & ChrW$(&H111) & ChrW$(&H1EA7) & "u):"
    btnSelectDest.Caption = "Ch" & ChrW$(&H1ECD) & "n..."
    
    lblCurrency.Caption = ChrW$(&H110) & ChrW$(&H1A1) & "n v" & ChrW$(&H1ECB) & " ti" & ChrW$(&H1EC1) & "n t" & ChrW$(&H1EC7) & ":"
    Dim prevCurrIdx As Long
    prevCurrIdx = cboCurrency.ListIndex
    cboCurrency.Clear
    cboCurrency.AddItem "Vi" & ChrW$(&H1EC7) & "t Nam " & ChrW$(&H110) & ChrW$(&H1ED3) & "ng (VND)"
    cboCurrency.AddItem ChrW$(&H110) & ChrW$(&HF4) & " la M" & ChrW$(&H1EF9) & " (USD)"
    cboCurrency.AddItem "Kh" & ChrW$(&HF4) & "ng " & ChrW$(&H111) & ChrW$(&H1A1) & "n v" & ChrW$(&H1ECB)
    If prevCurrIdx >= 0 Then
        cboCurrency.ListIndex = prevCurrIdx
    Else
        cboCurrency.ListIndex = 0
    End If
    
    lblCasing.Caption = "Ki" & ChrW$(&H1EC3) & "u ch" & ChrW$(&H1EEF) & ":"
    Dim prevCasingIdx As Long
    prevCasingIdx = cboCasing.ListIndex
    cboCasing.Clear
    cboCasing.AddItem "Vi" & ChrW$(&H1EBF) & "t hoa ch" & ChrW$(&H1EEF) & " " & ChrW$(&H111) & ChrW$(&H1EA7) & "u"
    cboCasing.AddItem "VI" & ChrW$(&H1EBE) & "T HOA TO" & ChrW$(&HC0) & "N B" & ChrW$(&H1ED8)
    cboCasing.AddItem "vi" & ChrW$(&H1EBF) & "t th" & ChrW$(&H1B0) & ChrW$(&H1EDD) & "ng to" & ChrW$(&HE0) & "n b" & ChrW$(&H1ED9)
    If prevCasingIdx >= 0 Then
        cboCasing.ListIndex = prevCasingIdx
    Else
        cboCasing.ListIndex = 0
    End If
    
    chkAddChan.Caption = "Th" & ChrW$(&HEA) & "m ch" & ChrW$(&H1EEF) & " '" & UnicodeText.WordChan() & "'"
    chkAddDong.Caption = "Th" & ChrW$(&HEA) & "m " & ChrW$(&H111) & ChrW$(&H1A1) & "n v" & ChrW$(&H1ECB) & " '" & UnicodeText.WordDong() & "'"
    chkAddPeriod.Caption = "Th" & ChrW$(&HEA) & "m d" & ChrW$(&H1EA5) & "u ch" & ChrW$(&H1EA5) & "m cu" & ChrW$(&H1ED1) & "i c" & ChrW$(&HE2) & "u"
    
    fraOutput.Caption = "Ch" & ChrW$(&H1EBF) & " " & ChrW$(&H111) & ChrW$(&H1ED9) & " xu" & ChrW$(&H1EA5) & "t k" & ChrW$(&H1EBF) & "t qu" & ChrW$(&H1EA3)
    optStatic.Caption = "V" & ChrW$(&H103) & "n b" & ChrW$(&H1EA3) & "n t" & ChrW$(&H129) & "nh (Static Text)"
    optFormula.Caption = "C" & ChrW$(&HF4) & "ng th" & ChrW$(&H1EE9) & "c (Formula)"
    
    lblCopyright.Caption = BuildInfo.APP_COPYRIGHT
    
    btnConvert.Caption = "Chuy" & ChrW$(&H1EC3) & "n " & ChrW$(&H111) & ChrW$(&H1ED5) & "i"
    btnClose.Caption = ChrW$(&H110) & ChrW$(&HF3) & "ng"
End Sub

Public Sub LoadDefaultState()
    Dim engOpts As VnEngineOptions
    Dim currOpts As VnCurrencyOptions
    Dim fmtOpts As VnFormatOptions
    Dim appOpts As VnAppSettings
    
    Settings.LoadAllSettings engOpts, currOpts, fmtOpts, appOpts
    
    cboCurrency.ListIndex = 0 ' VND
    If fmtOpts.Casing >= 0 And fmtOpts.Casing <= 2 Then
        cboCasing.ListIndex = fmtOpts.Casing
    Else
        cboCasing.ListIndex = 0
    End If
    
    chkAddChan.Value = currOpts.AddChan
    chkAddDong.Value = True
    chkAddPeriod.Value = fmtOpts.AddPeriod
    
    If appOpts.OutputMode = VnOutputFormula Then
        optFormula.Value = True
    Else
        optStatic.Value = True
    End If
    
    On Error Resume Next
    If TypeName(Selection) = "Range" Then
        Set mHostWorkbook = Selection.Parent.Parent
        Set mSourceRange = Selection
        txtSourceRange.Text = GetQualifiedAddress(Selection)
    End If
    On Error GoTo 0
    
    cboCurrency_Change
End Sub

Private Sub cboCurrency_Change()
    Select Case cboCurrency.ListIndex
        Case 0 ' VND
            chkAddDong.Enabled = True
            chkAddDong.Value = True
            chkAddDong.Caption = "Th" & ChrW$(&HEA) & "m " & ChrW$(&H111) & ChrW$(&H1A1) & "n v" & ChrW$(&H1ECB) & " '" & UnicodeText.WordDong() & "'"
            chkAddChan.Enabled = True
            chkAddChan.Caption = "Th" & ChrW$(&HEA) & "m ch" & ChrW$(&H1EEF) & " '" & UnicodeText.WordChan() & "'"
            optFormula.Enabled = True
            lblFormulaTip.Caption = ""
        Case 1 ' USD
            chkAddDong.Enabled = False
            chkAddDong.Value = False
            chkAddDong.Caption = "Kh" & ChrW$(&HF4) & "ng " & ChrW$(&HE1) & "p d" & ChrW$(&H1EE5) & "ng cho USD"
            chkAddChan.Enabled = True
            chkAddChan.Caption = "Th" & ChrW$(&HEA) & "m '" & UnicodeText.WordChan() & "' khi tr" & ChrW$(&HF2) & "n USD (kh" & ChrW$(&HF4) & "ng c" & ChrW$(&HF3) & " cent)"
            optFormula.Enabled = False
            optStatic.Value = True
            lblFormulaTip.Caption = "C" & ChrW$(&HF4) & "ng th" & ChrW$(&H1EE9) & "c kh" & ChrW$(&HF4) & "ng h" & ChrW$(&H1ED7) & " tr" & ChrW$(&H1EE3) & " " & ChrW$(&H111) & ChrW$(&HF4) & " la M" & ChrW$(&H1EF9) & ". T" & ChrW$(&H1EF1) & " " & ChrW$(&H111) & ChrW$(&H1ED9) & "ng chuy" & ChrW$(&H1EC3) & "n sang v" & ChrW$(&H103) & "n b" & ChrW$(&H1EA3) & "n t" & ChrW$(&H129) & "nh."
        Case 2 ' None
            chkAddDong.Enabled = False
            chkAddDong.Value = False
            chkAddDong.Caption = "Kh" & ChrW$(&HF4) & "ng c" & ChrW$(&HF3) & " " & ChrW$(&H111) & ChrW$(&H1A1) & "n v" & ChrW$(&H1ECB)
            chkAddChan.Enabled = False
            chkAddChan.Value = False
            chkAddChan.Caption = "Kh" & ChrW$(&HF4) & "ng " & ChrW$(&HE1) & "p d" & ChrW$(&H1EE5) & "ng"
            optFormula.Enabled = True
            lblFormulaTip.Caption = ""
    End Select
End Sub

Private Sub btnSelectSource_Click()
    Dim picked As Excel.Range
    Dim defaultAddr As String
    If Not mSourceRange Is Nothing Then
        defaultAddr = GetQualifiedAddress(mSourceRange)
    ElseIf TypeName(Selection) = "Range" Then
        defaultAddr = GetQualifiedAddress(Selection)
    Else
        defaultAddr = ""
    End If
    
    Me.Hide
    On Error Resume Next
    Set picked = Application.InputBox( _
        Prompt:="Ch" & ChrW$(&H1ECD) & "n v" & ChrW$(&HF9) & "ng d" & ChrW$(&H1EEF) & " li" & ChrW$(&H1EC7) & "u ngu" & ChrW$(&H1ED3) & "n c" & ChrW$(&H1EA7) & "n chuy" & ChrW$(&H1EC3) & "n " & ChrW$(&H111) & ChrW$(&H1ED5) & "i:", _
        Title:="Ch" & ChrW$(&H1ECD) & "n v" & ChrW$(&HF9) & "ng ngu" & ChrW$(&H1ED3) & "n", _
        Default:=defaultAddr, _
        Type:=8 _
    )
    On Error GoTo 0
    
    If Not picked Is Nothing Then
        If mHostWorkbook Is Nothing Then
            Set mHostWorkbook = picked.Parent.Parent
        End If
        If picked.Parent.Parent Is mHostWorkbook Then
            Set mSourceRange = picked
            txtSourceRange.Text = GetQualifiedAddress(picked)
        Else
            MsgBox "V" & ChrW$(&HF9) & "ng ch" & ChrW$(&H1ECD) & "n ph" & ChrW$(&H1EA3) & "i thu" & ChrW$(&H1ED9) & "c c" & ChrW$(&HF9) & "ng workbook hi" & ChrW$(&H1EC7) & "n t" & ChrW$(&H1EA1) & "i.", vbExclamation, BuildInfo.APP_NAME
        End If
    End If
    
    Me.Show
End Sub

Private Sub btnSelectDest_Click()
    Dim picked As Excel.Range
    Dim defaultAddr As String
    If Not mDestinationRange Is Nothing Then
        defaultAddr = GetQualifiedAddress(mDestinationRange)
    Else
        defaultAddr = ""
    End If
    
    Me.Hide
    On Error Resume Next
    Set picked = Application.InputBox( _
        Prompt:="Ch" & ChrW$(&H1ECD) & "n " & ChrW$(&HF4) & " ho" & ChrW$(&H1EA1) & "c v" & ChrW$(&HF9) & "ng xu" & ChrW$(&H1EA5) & "t k" & ChrW$(&H1EBF) & "t qu" & ChrW$(&H1EA3) & ":", _
        Title:="Ch" & ChrW$(&H1ECD) & "n v" & ChrW$(&HF9) & "ng " & ChrW$(&H111) & ChrW$(&HED) & "ch", _
        Default:=defaultAddr, _
        Type:=8 _
    )
    On Error GoTo 0
    
    If Not picked Is Nothing Then
        If mHostWorkbook Is Nothing Then
            Set mHostWorkbook = picked.Parent.Parent
        End If
        If picked.Parent.Parent Is mHostWorkbook Then
            Set mDestinationRange = picked
            txtDestRange.Text = GetQualifiedAddress(picked)
        Else
            MsgBox "V" & ChrW$(&HF9) & "ng ch" & ChrW$(&H1ECD) & "n ph" & ChrW$(&H1EA3) & "i thu" & ChrW$(&H1ED9) & "c c" & ChrW$(&HF9) & "ng workbook hi" & ChrW$(&H1EC7) & "n t" & ChrW$(&H1EA1) & "i.", vbExclamation, BuildInfo.APP_NAME
        End If
    End If
    
    Me.Show
End Sub

Private Sub txtSourceRange_AfterUpdate()
    On Error Resume Next
    If Len(Trim$(txtSourceRange.Text)) > 0 Then
        Dim resolved As Excel.Range
        Set resolved = ResolveRangeString(txtSourceRange.Text)
        If Not resolved Is Nothing Then
            Set mSourceRange = resolved
            If mHostWorkbook Is Nothing Then Set mHostWorkbook = resolved.Parent.Parent
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub txtDestRange_AfterUpdate()
    On Error Resume Next
    If Len(Trim$(txtDestRange.Text)) > 0 Then
        Dim resolved As Excel.Range
        Set resolved = ResolveRangeString(txtDestRange.Text)
        If Not resolved Is Nothing Then
            Set mDestinationRange = resolved
            If mHostWorkbook Is Nothing Then Set mHostWorkbook = resolved.Parent.Parent
        End If
    End If
    On Error GoTo 0
End Sub

Public Function ExecuteConversion(Optional ByVal SuppressPrompts As Boolean = False) As Boolean
    ExecuteConversion = False
    
    If mSourceRange Is Nothing And Len(Trim$(txtSourceRange.Text)) > 0 Then
        Set mSourceRange = ResolveRangeString(txtSourceRange.Text)
    End If
    If mDestinationRange Is Nothing And Len(Trim$(txtDestRange.Text)) > 0 Then
        Set mDestinationRange = ResolveRangeString(txtDestRange.Text)
    End If
    
    If mSourceRange Is Nothing Then
        If Not SuppressPrompts Then
            MsgBox "Vui l" & ChrW$(&HF2) & "ng ch" & ChrW$(&H1ECD) & "n v" & ChrW$(&HF9) & "ng d" & ChrW$(&H1EEF) & " li" & ChrW$(&H1EC7) & "u ngu" & ChrW$(&H1ED3) & "n.", vbExclamation, BuildInfo.APP_NAME
        End If
        Exit Function
    End If
    
    If mDestinationRange Is Nothing Then
        If Not SuppressPrompts Then
            MsgBox "Vui l" & ChrW$(&HF2) & "ng ch" & ChrW$(&H1ECD) & "n v" & ChrW$(&HF9) & "ng xu" & ChrW$(&H1EA5) & "t k" & ChrW$(&H1EBF) & "t qu" & ChrW$(&H1EA3) & ".", vbExclamation, BuildInfo.APP_NAME
        End If
        Exit Function
    End If
    
    If mSourceRange.Areas.Count <> 1 Or mDestinationRange.Areas.Count <> 1 Then
        If Not SuppressPrompts Then
            MsgBox "Kh" & ChrW$(&HF4) & "ng h" & ChrW$(&H1ED7) & " tr" & ChrW$(&H1EE3) & " v" & ChrW$(&HF9) & "ng ch" & ChrW$(&H1ECD) & "n kh" & ChrW$(&HF4) & "ng li" & ChrW$(&HEA) & "n t" & ChrW$(&H1EE5) & "c (multi-area).", vbExclamation, BuildInfo.APP_NAME
        End If
        Exit Function
    End If
    
    If Not (mSourceRange.Parent.Parent Is mDestinationRange.Parent.Parent) Then
        If Not SuppressPrompts Then
            MsgBox "V" & ChrW$(&HF9) & "ng ngu" & ChrW$(&H1ED3) & "n v" & ChrW$(&HE0) & " v" & ChrW$(&HF9) & "ng " & ChrW$(&H111) & ChrW$(&HED) & "ch ph" & ChrW$(&H1EA3) & "i thu" & ChrW$(&H1ED9) & "c c" & ChrW$(&HF9) & "ng m" & ChrW$(&H1ED9) & "t workbook.", vbExclamation, BuildInfo.APP_NAME
        End If
        Exit Function
    End If
    
    If mSourceRange.Parent Is mDestinationRange.Parent Then
        Dim destExp As Excel.Range
        If mDestinationRange.Cells.Count = 1 Then
            Set destExp = mDestinationRange.Resize(mSourceRange.Rows.Count, mSourceRange.Columns.Count)
        Else
            Set destExp = mDestinationRange
        End If
        If Not Application.Intersect(mSourceRange, destExp) Is Nothing Then
            If Not SuppressPrompts Then
                MsgBox "V" & ChrW$(&HF9) & "ng ngu" & ChrW$(&H1ED3) & "n v" & ChrW$(&HE0) & " v" & ChrW$(&HF9) & "ng " & ChrW$(&H111) & ChrW$(&HED) & "ch kh" & ChrW$(&HF4) & "ng " & ChrW$(&H111) & ChrW$(&H1B0) & ChrW$(&H1EE3) & "c " & ChrW$(&H111) & ChrW$(&HE8) & " l" & ChrW$(&HEA) & "n nhau tr" & ChrW$(&HEA) & "n c" & ChrW$(&HF9) & "ng m" & ChrW$(&H1ED9) & "t sheet.", vbCritical, BuildInfo.APP_NAME
            End If
            Exit Function
        End If
    End If
    
    If mDestinationRange.Parent.ProtectContents Then
        Dim destCheckRange As Excel.Range
        If mDestinationRange.Cells.Count = 1 Then
            Set destCheckRange = mDestinationRange.Resize(mSourceRange.Rows.Count, mSourceRange.Columns.Count)
        Else
            Set destCheckRange = mDestinationRange
        End If
        Dim checkCell As Excel.Range
        For Each checkCell In destCheckRange.Cells
            If checkCell.Locked Then
                If Not SuppressPrompts Then
                    MsgBox "V" & ChrW$(&HF9) & "ng " & ChrW$(&H111) & ChrW$(&HED) & "ch ch" & ChrW$(&H1EE9) & "a " & ChrW$(&HF4) & " b" & ChrW$(&H1ECB) & " kh" & ChrW$(&HF3) & "a tr" & ChrW$(&HEA) & "n sheet " & ChrW$(&H111) & ChrW$(&H1B0) & ChrW$(&H1EE3) & "c b" & ChrW$(&H1EA3) & "o v" & ChrW$(&H1EC7) & ".", vbCritical, BuildInfo.APP_NAME
                End If
                Exit Function
            End If
        Next checkCell
    End If
    
    Dim engOpts As VnEngineOptions
    Dim currOpts As VnCurrencyOptions
    Dim fmtOpts As VnFormatOptions
    Dim appOpts As VnAppSettings
    Settings.LoadAllSettings engOpts, currOpts, fmtOpts, appOpts
    
    Select Case cboCurrency.ListIndex
        Case 0 ' VND
            If chkAddDong.Value Then
                currOpts = DefaultCurrencyOptions(VnCurrVND)
            Else
                currOpts = DefaultCurrencyOptions(VnCurrNone)
            End If
            currOpts.AddChan = chkAddChan.Value
        Case 1 ' USD
            currOpts = DefaultCurrencyOptions(VnCurrUSD)
            currOpts.AddChan = chkAddChan.Value
        Case 2 ' None
            currOpts = DefaultCurrencyOptions(VnCurrNone)
            currOpts.AddChan = False
    End Select
    
    If cboCasing.ListIndex >= 0 Then
        fmtOpts.Casing = cboCasing.ListIndex
    End If
    fmtOpts.AddPeriod = chkAddPeriod.Value
    
    If optFormula.Value And optFormula.Enabled Then
        appOpts.OutputMode = VnOutputFormula
    Else
        appOpts.OutputMode = VnOutputStatic
    End If
    
    If appOpts.ConfirmOverwrite And Not SuppressPrompts Then
        Dim targetArea As Excel.Range
        If mDestinationRange.Cells.Count = 1 Then
            Set targetArea = mDestinationRange.Resize(mSourceRange.Rows.Count, mSourceRange.Columns.Count)
        Else
            Set targetArea = mDestinationRange
        End If
        If Application.WorksheetFunction.CountA(targetArea) > 0 Then
            Dim ans As VbMsgBoxResult
            ans = MsgBox("V" & ChrW$(&HF9) & "ng " & ChrW$(&H111) & ChrW$(&HED) & "ch " & ChrW$(&H111) & ChrW$(&HE3) & " c" & ChrW$(&HF3) & " d" & ChrW$(&H1EEF) & " li" & ChrW$(&H1EC7) & "u. B" & ChrW$(&H1EA1) & "n c" & ChrW$(&HF3) & " ch" & ChrW$(&H1EAF) & "c ch" & ChrW$(&H1EAF) & "n mu" & ChrW$(&H1ED1) & "n ghi " & ChrW$(&H111) & ChrW$(&HE8) & " kh" & ChrW$(&HF4) & "ng?", vbQuestion + vbYesNo + vbDefaultButton2, BuildInfo.APP_NAME)
            If ans <> vbYes Then Exit Function
        End If
    End If
    
    mBatchResult = CellProcessor.ConvertRange(mSourceRange, mDestinationRange, engOpts, currOpts, fmtOpts, appOpts)
    
    If Not mBatchResult.Success Then
        If Not SuppressPrompts Then
            MsgBox mBatchResult.ErrorMessage, vbExclamation, BuildInfo.APP_NAME
        End If
        Exit Function
    End If
    
    On Error Resume Next
    Application.Run "InvalidateUndoRibbon"
    On Error GoTo 0
    
    If appOpts.ShowBatchSummary And mBatchResult.convertedCount > 1 And Not SuppressPrompts Then
        MsgBox ChrW$(&H110) & ChrW$(&HE3) & " chuy" & ChrW$(&H1EC3) & "n " & ChrW$(&H111) & ChrW$(&H1ED5) & "i th" & ChrW$(&HE0) & "nh c" & ChrW$(&HF4) & "ng " & mBatchResult.convertedCount & " " & ChrW$(&HF4) & "." & _
               IIf(mBatchResult.skippedCount > 0, vbCrLf & "B" & ChrW$(&H1ECF) & " qua: " & mBatchResult.skippedCount & " " & ChrW$(&HF4) & ".", ""), _
               vbInformation, BuildInfo.APP_NAME
    End If
    
    ExecuteConversion = True
End Function

Private Function GetQualifiedAddress(ByVal rng As Excel.Range) As String
    If rng Is Nothing Then
        GetQualifiedAddress = ""
    Else
        GetQualifiedAddress = "'" & rng.Parent.Name & "'!" & rng.Address(True, True)
    End If
End Function

Private Function ResolveRangeString(ByVal addr As String) As Excel.Range
    On Error Resume Next
    Dim posExcl As Long
    posExcl = InStr(addr, "!")
    If posExcl > 0 Then
        Dim sName As String
        sName = Left$(addr, posExcl - 1)
        If Left$(sName, 1) = "'" And Right$(sName, 1) = "'" Then
            sName = Mid$(sName, 2, Len(sName) - 2)
        End If
        Dim rPart As String
        rPart = Mid$(addr, posExcl + 1)
        If Not mHostWorkbook Is Nothing Then
            Set ResolveRangeString = mHostWorkbook.Worksheets(sName).Range(rPart)
        Else
            Set ResolveRangeString = Application.Range(addr)
        End If
    Else
        If Not mHostWorkbook Is Nothing Then
            Set ResolveRangeString = mHostWorkbook.ActiveSheet.Range(addr)
        Else
            Set ResolveRangeString = Range(addr)
        End If
    End If
    On Error GoTo 0
End Function


