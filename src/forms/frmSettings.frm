VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "UserForm1"
   ClientHeight    =   9420.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7365
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' frmSettings.frm
' Settings configuration dialog for BWPConvertTTNVN.
' Pure 7-bit ASCII safe source code.
' All Vietnamese captions assigned dynamically at runtime via ChrW$().
' ==============================================================================

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnReset_Click()
    ' Loads factory defaults into UI controls only; DOES NOT modify Registry!
    LoadDefaultsToUI
End Sub

Private Sub btnSave_Click()
    ' Validates and persists current UI state to Registry, then closes.
    SaveUIToSettings
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    PopulateCaptions
    LoadFromSettings
End Sub

Public Sub PopulateCaptions()
    ' Title: CAI DAT
    Me.Caption = "C" & ChrW$(&HC0) & "I " & ChrW$(&H110) & ChrW$(&H1EB6) & "T"
    
    ' Grammar frame
    fraGrammar.Caption = "Quy t" & ChrW$(&H1EAF) & "c ng" & ChrW$(&H1EEF) & " ph" & ChrW$(&HE1) & "p"
    lblZero.Caption = ChrW$(&H110) & ChrW$(&H1ECD) & "c s" & ChrW$(&H1ED0) & " 0 " & ChrW$(&H1A1) & " h" & ChrW$(&HE0) & "ng ch" & ChrW$(&H1EE5) & "c:"
    optZeroLe.Caption = UnicodeText.WordLe()
    optZeroLinh.Caption = UnicodeText.WordLinh()
    
    lblThousand.Caption = "H" & ChrW$(&HE0) & "ng ngh" & ChrW$(&HEC) & "n:"
    optThousandNghin.Caption = UnicodeText.WordNghin()
    optThousandNgan.Caption = UnicodeText.WordNgan()
    
    lblFour.Caption = "S" & ChrW$(&H1ED1) & " 4 trong 24..94:"
    optFourTu.Caption = UnicodeText.WordTu() & " (hai " & UnicodeText.WordMuoiChuc() & " " & UnicodeText.WordTu() & ")"
    optFourBon.Caption = UnicodeText.WordBon() & " (hai " & UnicodeText.WordMuoiChuc() & " " & UnicodeText.WordBon() & ")"
    
    ' Formatting defaults frame
    fraFormat.Caption = ChrW$(&H110) & ChrW$(&H1ECB) & "nh d" & ChrW$(&H1EA1) & "ng m" & ChrW$(&H1EB7) & "c " & ChrW$(&H111) & ChrW$(&H1ECB) & "nh"
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
    chkAddPeriod.Caption = "Th" & ChrW$(&HEA) & "m d" & ChrW$(&H1EA5) & "u ch" & ChrW$(&H1EA5) & "m cu" & ChrW$(&H1ED1) & "i c" & ChrW$(&HE2) & "u"
    chkAddDong.Caption = "Th" & ChrW$(&HEA) & "m " & ChrW$(&H111) & ChrW$(&H1A1) & "n v" & ChrW$(&H1ECB) & " '" & UnicodeText.WordDong() & "'"
    
    ' Output mode frame
    fraOutput.Caption = "Ch" & ChrW$(&H1EBF) & " " & ChrW$(&H111) & ChrW$(&H1ED9) & " xu" & ChrW$(&H1EA5) & "t k" & ChrW$(&H1EBF) & "t qu" & ChrW$(&H1EA3)
    optOutputStatic.Caption = "V" & ChrW$(&H103) & "n b" & ChrW$(&H1EA3) & "n t" & ChrW$(&H129) & "nh (Static Text)"
    optOutputFormula.Caption = "C" & ChrW$(&HF4) & "ng th" & ChrW$(&H1EE9) & "c (Formula)"
    
    ' Quick convert frame
    fraQuick.Caption = "H" & ChrW$(&H1B0) & ChrW$(&H1EDB) & "ng chuy" & ChrW$(&H1EC3) & "n " & ChrW$(&H111) & ChrW$(&H1ED5) & "i nhanh"
    optQuickAuto.Caption = "T" & ChrW$(&H1EF1) & " " & ChrW$(&H111) & ChrW$(&H1ED9) & "ng (C" & ChrW$(&H1ED9) & "t->Ph" & ChrW$(&H1EA3) & "i, D" & ChrW$(&HF2) & "ng->D" & ChrW$(&H1B0) & ChrW$(&H1EDB) & "i)"
    optQuickRight.Caption = "Lu" & ChrW$(&HF4) & "n sang ph" & ChrW$(&H1EA3) & "i"
    optQuickBelow.Caption = "Lu" & ChrW$(&HF4) & "n xu" & ChrW$(&H1ED1) & "ng d" & ChrW$(&H1B0) & ChrW$(&H1EDB) & "i"
    
    ' Safety frame
    fraSafety.Caption = "An to" & ChrW$(&HE0) & "n d" & ChrW$(&H1EEF) & " li" & ChrW$(&H1EC7) & "u"
    chkConfirmOverwrite.Caption = "X" & ChrW$(&HE1) & "c nh" & ChrW$(&H1EAD) & "n tr" & ChrW$(&H1B0) & ChrW$(&H1EDB) & "c khi ghi " & ChrW$(&H111) & ChrW$(&HE8) & " d" & ChrW$(&H1EEF) & " li" & ChrW$(&H1EC7) & "u c" & ChrW$(&HF3) & " s" & ChrW$(&H1EB5) & "n"
    chkShowBatchSummary.Caption = "Hi" & ChrW$(&H1EC3) & "n th" & ChrW$(&H1ECB) & " t" & ChrW$(&H1ED5) & "ng k" & ChrW$(&H1EBF) & "t sau khi chuy" & ChrW$(&H1EC3) & "n " & ChrW$(&H111) & ChrW$(&H1ED5) & "i h" & ChrW$(&HE0) & "ng lo" & ChrW$(&H1EA1) & "t"
    
    ' Action buttons
    btnReset.Caption = "M" & ChrW$(&H1EB7) & "c " & ChrW$(&H111) & ChrW$(&H1ECB) & "nh"
    btnSave.Caption = "L" & ChrW$(&H1B0) & "u"
    btnCancel.Caption = "H" & ChrW$(&H1EE7) & "y"
End Sub

Public Sub LoadDefaultsToUI()
    Dim engOpts As VnEngineOptions
    Dim currOpts As VnCurrencyOptions
    Dim fmtOpts As VnFormatOptions
    Dim appOpts As VnAppSettings
    
    engOpts = DefaultEngineOptions()
    currOpts = DefaultCurrencyOptions(VnCurrVND)
    fmtOpts = DefaultFormatOptions()
    appOpts = DefaultAppSettings()
    
    ApplyToUI engOpts, currOpts, fmtOpts, appOpts
End Sub

Public Sub LoadFromSettings()
    Dim engOpts As VnEngineOptions
    Dim currOpts As VnCurrencyOptions
    Dim fmtOpts As VnFormatOptions
    Dim appOpts As VnAppSettings
    
    Settings.LoadAllSettings engOpts, currOpts, fmtOpts, appOpts
    ApplyToUI engOpts, currOpts, fmtOpts, appOpts
End Sub

Private Sub ApplyToUI( _
    ByRef engOpts As VnEngineOptions, _
    ByRef currOpts As VnCurrencyOptions, _
    ByRef fmtOpts As VnFormatOptions, _
    ByRef appOpts As VnAppSettings _
)
    If engOpts.ZeroStyle = VnZeroLinh Then
        optZeroLinh.Value = True
    Else
        optZeroLe.Value = True
    End If
    
    If engOpts.ThousandStyle = VnThousandNgan Then
        optThousandNgan.Value = True
    Else
        optThousandNghin.Value = True
    End If
    
    If engOpts.FourStyle = VnFourBon Then
        optFourBon.Value = True
    Else
        optFourTu.Value = True
    End If
    
    If fmtOpts.Casing >= 0 And fmtOpts.Casing <= 2 Then
        cboCasing.ListIndex = fmtOpts.Casing
    Else
        cboCasing.ListIndex = 0
    End If
    
    chkAddChan.Value = currOpts.AddChan
    chkAddPeriod.Value = fmtOpts.AddPeriod
    chkAddDong.Value = (currOpts.CurrencyType = VnCurrVND)
    
    If appOpts.OutputMode = VnOutputFormula Then
        optOutputFormula.Value = True
    Else
        optOutputStatic.Value = True
    End If
    
    Select Case appOpts.QuickConvertDirection
        Case VnQuickRight
            optQuickRight.Value = True
        Case VnQuickBelow
            optQuickBelow.Value = True
        Case Else
            optQuickAuto.Value = True
    End Select
    
    chkConfirmOverwrite.Value = appOpts.ConfirmOverwrite
    chkShowBatchSummary.Value = appOpts.ShowBatchSummary
End Sub

Public Sub SaveUIToSettings()
    Dim engOpts As VnEngineOptions
    Dim currOpts As VnCurrencyOptions
    Dim fmtOpts As VnFormatOptions
    Dim appOpts As VnAppSettings
    
    Settings.LoadAllSettings engOpts, currOpts, fmtOpts, appOpts
    
    If optZeroLinh.Value Then
        engOpts.ZeroStyle = VnZeroLinh
    Else
        engOpts.ZeroStyle = VnZeroLe
    End If
    
    If optThousandNgan.Value Then
        engOpts.ThousandStyle = VnThousandNgan
    Else
        engOpts.ThousandStyle = VnThousandNghin
    End If
    
    If optFourBon.Value Then
        engOpts.FourStyle = VnFourBon
    Else
        engOpts.FourStyle = VnFourTu
    End If
    
    If chkAddDong.Value Then
        currOpts.CurrencyType = VnCurrVND
        currOpts.CustomSuffix = UnicodeText.WordDong()
    Else
        currOpts.CurrencyType = VnCurrNone
        currOpts.CustomSuffix = ""
    End If
    currOpts.AddChan = chkAddChan.Value
    
    If cboCasing.ListIndex >= 0 Then
        fmtOpts.Casing = cboCasing.ListIndex
    End If
    fmtOpts.AddPeriod = chkAddPeriod.Value
    
    If optOutputFormula.Value Then
        appOpts.OutputMode = VnOutputFormula
    Else
        appOpts.OutputMode = VnOutputStatic
    End If
    
    If optQuickRight.Value Then
        appOpts.QuickConvertDirection = VnQuickRight
    ElseIf optQuickBelow.Value Then
        appOpts.QuickConvertDirection = VnQuickBelow
    Else
        appOpts.QuickConvertDirection = VnQuickAuto
    End If
    
    appOpts.ConfirmOverwrite = chkConfirmOverwrite.Value
    appOpts.ShowBatchSummary = chkShowBatchSummary.Value
    appOpts.SettingsVersion = 1
    
    Settings.SaveAllSettings engOpts, currOpts, fmtOpts, appOpts
End Sub

Public Property Get ZeroStyleValue() As Long
    If optZeroLinh.Value Then ZeroStyleValue = 1 Else ZeroStyleValue = 0
End Property

Public Property Get TitleText() As String
    TitleText = "C" & ChrW$(&H00C0) & "I " & ChrW$(&H0110) & ChrW$(&H1EB6) & "T"
End Property

Public Property Let ZeroStyleValue(ByVal val As Long)
    If val = 1 Then optZeroLinh.Value = True Else optZeroLe.Value = True
End Property

Public Property Get ThousandStyleValue() As Long
    If optThousandNgan.Value Then ThousandStyleValue = 1 Else ThousandStyleValue = 0
End Property

Public Property Let ThousandStyleValue(ByVal val As Long)
    If val = 1 Then optThousandNgan.Value = True Else optThousandNghin.Value = True
End Property

Public Property Get FourStyleValue() As Long
    If optFourBon.Value Then FourStyleValue = 1 Else FourStyleValue = 0
End Property

Public Property Let FourStyleValue(ByVal val As Long)
    If val = 1 Then optFourBon.Value = True Else optFourTu.Value = True
End Property

Public Property Get CasingIndex() As Long
    CasingIndex = cboCasing.ListIndex
End Property

Public Property Let CasingIndex(ByVal val As Long)
    cboCasing.ListIndex = val
End Property

Public Property Get AddChanValue() As Boolean
    AddChanValue = chkAddChan.Value
End Property

Public Property Let AddChanValue(ByVal val As Boolean)
    chkAddChan.Value = val
End Property

Public Property Get AddPeriodValue() As Boolean
    AddPeriodValue = chkAddPeriod.Value
End Property

Public Property Let AddPeriodValue(ByVal val As Boolean)
    chkAddPeriod.Value = val
End Property

Public Property Get AddDongValue() As Boolean
    AddDongValue = chkAddDong.Value
End Property

Public Property Let AddDongValue(ByVal val As Boolean)
    chkAddDong.Value = val
End Property

Public Property Get OutputModeValue() As Long
    If optOutputFormula.Value Then OutputModeValue = 1 Else OutputModeValue = 0
End Property

Public Property Let OutputModeValue(ByVal val As Long)
    If val = 1 Then optOutputFormula.Value = True Else optOutputStatic.Value = True
End Property

Public Property Get QuickDirectionValue() As Long
    If optQuickRight.Value Then
        QuickDirectionValue = 1
    ElseIf optQuickBelow.Value Then
        QuickDirectionValue = 2
    Else
        QuickDirectionValue = 0
    End If
End Property

Public Property Let QuickDirectionValue(ByVal val As Long)
    Select Case val
        Case 1: optQuickRight.Value = True
        Case 2: optQuickBelow.Value = True
        Case Else: optQuickAuto.Value = True
    End Select
End Property

Public Property Get ConfirmOverwrite() As Boolean
    ConfirmOverwrite = chkConfirmOverwrite.Value
End Property

Public Property Let ConfirmOverwrite(ByVal val As Boolean)
    chkConfirmOverwrite.Value = val
End Property

Public Property Get ShowBatchSummary() As Boolean
    ShowBatchSummary = chkShowBatchSummary.Value
End Property

Public Property Let ShowBatchSummary(ByVal val As Boolean)
    chkShowBatchSummary.Value = val
End Property

Public Sub TriggerReset()
    btnReset_Click
End Sub

Public Sub TriggerSave()
    btnSave_Click
End Sub

Public Sub TriggerCancel()
    btnCancel_Click
End Sub


