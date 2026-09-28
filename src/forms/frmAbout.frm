VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "UserForm1"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6165
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' frmAbout.frm
' About dialog for BWPConvertTTNVN.
' Pure 7-bit ASCII safe source code.
' All Vietnamese captions assigned dynamically at runtime via ChrW$() / BuildInfo.
' ==============================================================================

Private Sub btnOK_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    PopulateCaptions
End Sub

Public Sub PopulateCaptions()
    ' Title: About
    Me.Caption = "About"
    
    lblAppName.Caption = BuildInfo.APP_NAME
    lblVersion.Caption = "Phi" & ChrW$(&HEA) & "n b" & ChrW$(&H1EA3) & "n " & BuildInfo.APP_VERSION
    
    ' Mandatory copyright notice: Copyright (C) 2026 - IT Leon
    lblCopyright.Caption = BuildInfo.APP_COPYRIGHT
    
    ' Bullet points
    lblBullets.Caption = ChrW$(&H2022) & " Ho" & ChrW$(&H1EA1) & "t " & ChrW$(&H111) & ChrW$(&H1ED9) & "ng 100% Offline, kh" & ChrW$(&HF4) & "ng y" & ChrW$(&HEA) & "u c" & ChrW$(&H1EA7) & "u k" & ChrW$(&H1EBF) & "t n" & ChrW$(&H1ED1) & "i m" & ChrW$(&H1EA1) & "ng." & vbCrLf & _
                         ChrW$(&H2022) & " Kh" & ChrW$(&HF4) & "ng thu th" & ChrW$(&H1EAD) & "p d" & ChrW$(&H1EEF) & " li" & ChrW$(&H1EC7) & "u ho" & ChrW$(&H1EA1) & "c telemetry." & vbCrLf & _
                         ChrW$(&H2022) & " Kh" & ChrW$(&HF4) & "ng y" & ChrW$(&HEA) & "u c" & ChrW$(&H1EA7) & "u DLL b" & ChrW$(&HEA) & "n th" & ChrW$(&H1EE9) & " ba." & vbCrLf & _
                         ChrW$(&H2022) & " M" & ChrW$(&HE3) & " ngu" & ChrW$(&H1ED3) & "n m" & ChrW$(&H1EDF) & " (MIT License), c" & ChrW$(&HF3) & " th" & ChrW$(&H1EC3) & " ki" & ChrW$(&H1EC3) & "m tra v" & ChrW$(&HE0) & " audit."
                         
    btnOK.Caption = ChrW$(&H110) & ChrW$(&HF3) & "ng"
End Sub

Public Property Get TitleText() As String
    TitleText = "About"
End Property

Public Property Get AppNameText() As String
    AppNameText = lblAppName.Caption
End Property

Public Property Get VersionText() As String
    VersionText = lblVersion.Caption
End Property

Public Property Get CopyrightText() As String
    CopyrightText = lblCopyright.Caption
End Property

Public Property Get BulletsText() As String
    BulletsText = lblBullets.Caption
End Property

Public Property Get CloseButtonCaption() As String
    CloseButtonCaption = btnOK.Caption
End Property


