Attribute VB_Name = "BuildInfo"
Option Explicit

Public Const APP_NAME As String = "BWPConvertTTNVN"
Public Const APP_VERSION As String = "1.0.0"

Public Property Get APP_COPYRIGHT() As String
    APP_COPYRIGHT = "Copyright " & ChrW$(&HA9) & " 2026 - IT Leon"
End Property
