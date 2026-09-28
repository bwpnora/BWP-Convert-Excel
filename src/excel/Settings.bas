Attribute VB_Name = "Settings"
Option Explicit

' ==============================================================================
' Settings.bas
' Persistent registry configuration management for BWPConvertTTNVN.
' Registry base: HKCU\Software\VB and VBA Program Settings\BWPConvertTTNVN
' Pure 7-bit ASCII safe source code.
' ==============================================================================

' Output Mode Enum
Public Enum VnOutputMode
    VnOutputStatic = 0
    VnOutputFormula = 1
End Enum

' Quick Convert Direction Enum
Public Enum VnQuickDirection
    VnQuickAuto = 0
    VnQuickRight = 1
    VnQuickBelow = 2
End Enum

' Application Settings User-Defined Type (UDT)
Public Type VnAppSettings
    OutputMode As VnOutputMode
    QuickConvertDirection As VnQuickDirection
    ConfirmOverwrite As Boolean
    ShowBatchSummary As Boolean
    SettingsVersion As Long
End Type

' Registry Constants
Private Const REG_APP_NAME As String = "BWPConvertTTNVN"
Private Const REG_SEC_ENGINE As String = "Engine"
Private Const REG_SEC_CURRENCY As String = "Currency"
Private Const REG_SEC_FORMAT As String = "Format"
Private Const REG_SEC_APP As String = "App"

' ==============================================================================
' Factory Function for Default App Settings
' ==============================================================================

Public Function DefaultAppSettings() As VnAppSettings
    Dim Opts As VnAppSettings
    Opts.OutputMode = VnOutputStatic
    Opts.QuickConvertDirection = VnQuickAuto
    Opts.ConfirmOverwrite = True
    Opts.ShowBatchSummary = True
    Opts.SettingsVersion = 1
    DefaultAppSettings = Opts
End Function

' ==============================================================================
' Settings Persistence: Load, Save, Reset
' ==============================================================================

Public Sub LoadAllSettings( _
    ByRef EngineOpts As VnEngineOptions, _
    ByRef CurrOpts As VnCurrencyOptions, _
    ByRef FormatOpts As VnFormatOptions, _
    ByRef AppOpts As VnAppSettings _
)
    Dim defEng As VnEngineOptions
    Dim defCurr As VnCurrencyOptions
    Dim defFmt As VnFormatOptions
    Dim defApp As VnAppSettings
    
    defEng = DefaultEngineOptions()
    defCurr = DefaultCurrencyOptions(VnCurrVND)
    defFmt = DefaultFormatOptions()
    defApp = DefaultAppSettings()
    
    On Error Resume Next
    
    ' 1. Engine Options
    EngineOpts.ZeroStyle = ParseLong(GetSetting(REG_APP_NAME, REG_SEC_ENGINE, "ZeroStyle", ""), 0, 1, defEng.ZeroStyle)
    EngineOpts.ThousandStyle = ParseLong(GetSetting(REG_APP_NAME, REG_SEC_ENGINE, "ThousandStyle", ""), 0, 1, defEng.ThousandStyle)
    EngineOpts.FourStyle = ParseLong(GetSetting(REG_APP_NAME, REG_SEC_ENGINE, "FourStyle", ""), 0, 1, defEng.FourStyle)
    EngineOpts.DecimalMode = ParseLong(GetSetting(REG_APP_NAME, REG_SEC_ENGINE, "DecimalMode", ""), 0, 1, defEng.DecimalMode)
    
    ' 2. Currency Options
    CurrOpts.CurrencyType = ParseLong(GetSetting(REG_APP_NAME, REG_SEC_CURRENCY, "CurrencyType", ""), 0, 3, defCurr.CurrencyType)
    CurrOpts.AddChan = ParseBool(GetSetting(REG_APP_NAME, REG_SEC_CURRENCY, "AddChan", ""), defCurr.AddChan)
    CurrOpts.DecimalPlaces = CInt(ParseLong(GetSetting(REG_APP_NAME, REG_SEC_CURRENCY, "DecimalPlaces", ""), 0, 4, defCurr.DecimalPlaces))
    CurrOpts.CustomPrefix = GetSetting(REG_APP_NAME, REG_SEC_CURRENCY, "CustomPrefix", defCurr.CustomPrefix)
    CurrOpts.CustomSuffix = GetSetting(REG_APP_NAME, REG_SEC_CURRENCY, "CustomSuffix", defCurr.CustomSuffix)
    CurrOpts.CustomSubUnit = GetSetting(REG_APP_NAME, REG_SEC_CURRENCY, "CustomSubUnit", defCurr.CustomSubUnit)
    
    ' 3. Format Options
    FormatOpts.Casing = ParseLong(GetSetting(REG_APP_NAME, REG_SEC_FORMAT, "Casing", ""), 0, 2, defFmt.Casing)
    FormatOpts.AddPeriod = ParseBool(GetSetting(REG_APP_NAME, REG_SEC_FORMAT, "AddPeriod", ""), defFmt.AddPeriod)
    
    ' 4. Application Options
    AppOpts.OutputMode = ParseLong(GetSetting(REG_APP_NAME, REG_SEC_APP, "OutputMode", ""), 0, 1, defApp.OutputMode)
    AppOpts.QuickConvertDirection = ParseLong(GetSetting(REG_APP_NAME, REG_SEC_APP, "QuickConvertDirection", ""), 0, 2, defApp.QuickConvertDirection)
    AppOpts.ConfirmOverwrite = ParseBool(GetSetting(REG_APP_NAME, REG_SEC_APP, "ConfirmOverwrite", ""), defApp.ConfirmOverwrite)
    AppOpts.ShowBatchSummary = ParseBool(GetSetting(REG_APP_NAME, REG_SEC_APP, "ShowBatchSummary", ""), defApp.ShowBatchSummary)
    AppOpts.SettingsVersion = ParseLong(GetSetting(REG_APP_NAME, REG_SEC_APP, "SettingsVersion", ""), 1, 100, defApp.SettingsVersion)
    
    On Error GoTo 0
End Sub

Public Sub SaveAllSettings( _
    ByRef EngineOpts As VnEngineOptions, _
    ByRef CurrOpts As VnCurrencyOptions, _
    ByRef FormatOpts As VnFormatOptions, _
    ByRef AppOpts As VnAppSettings _
)
    On Error Resume Next
    
    ' 1. Engine Options
    SaveSetting REG_APP_NAME, REG_SEC_ENGINE, "ZeroStyle", CStr(CLng(EngineOpts.ZeroStyle))
    SaveSetting REG_APP_NAME, REG_SEC_ENGINE, "ThousandStyle", CStr(CLng(EngineOpts.ThousandStyle))
    SaveSetting REG_APP_NAME, REG_SEC_ENGINE, "FourStyle", CStr(CLng(EngineOpts.FourStyle))
    SaveSetting REG_APP_NAME, REG_SEC_ENGINE, "DecimalMode", CStr(CLng(EngineOpts.DecimalMode))
    
    ' 2. Currency Options
    SaveSetting REG_APP_NAME, REG_SEC_CURRENCY, "CurrencyType", CStr(CLng(CurrOpts.CurrencyType))
    SaveSetting REG_APP_NAME, REG_SEC_CURRENCY, "AddChan", IIf(CurrOpts.AddChan, "1", "0")
    SaveSetting REG_APP_NAME, REG_SEC_CURRENCY, "DecimalPlaces", CStr(CLng(CurrOpts.DecimalPlaces))
    SaveSetting REG_APP_NAME, REG_SEC_CURRENCY, "CustomPrefix", CurrOpts.CustomPrefix
    SaveSetting REG_APP_NAME, REG_SEC_CURRENCY, "CustomSuffix", CurrOpts.CustomSuffix
    SaveSetting REG_APP_NAME, REG_SEC_CURRENCY, "CustomSubUnit", CurrOpts.CustomSubUnit
    
    ' 3. Format Options
    SaveSetting REG_APP_NAME, REG_SEC_FORMAT, "Casing", CStr(CLng(FormatOpts.Casing))
    SaveSetting REG_APP_NAME, REG_SEC_FORMAT, "AddPeriod", IIf(FormatOpts.AddPeriod, "1", "0")
    
    ' 4. Application Options
    SaveSetting REG_APP_NAME, REG_SEC_APP, "OutputMode", CStr(CLng(AppOpts.OutputMode))
    SaveSetting REG_APP_NAME, REG_SEC_APP, "QuickConvertDirection", CStr(CLng(AppOpts.QuickConvertDirection))
    SaveSetting REG_APP_NAME, REG_SEC_APP, "ConfirmOverwrite", IIf(AppOpts.ConfirmOverwrite, "1", "0")
    SaveSetting REG_APP_NAME, REG_SEC_APP, "ShowBatchSummary", IIf(AppOpts.ShowBatchSummary, "1", "0")
    SaveSetting REG_APP_NAME, REG_SEC_APP, "SettingsVersion", "1"
    
    On Error GoTo 0
End Sub

Public Sub ResetSettingsToDefault()
    On Error Resume Next
    DeleteSetting REG_APP_NAME
    On Error GoTo 0
End Sub

' ==============================================================================
' Internal Validation & Parsing Helpers
' ==============================================================================

Private Function ParseLong( _
    ByVal ValStr As String, _
    ByVal MinVal As Long, _
    ByVal MaxVal As Long, _
    ByVal DefaultVal As Long _
) As Long
    On Error GoTo Fallback
    If Len(Trim$(ValStr)) = 0 Then
        ParseLong = DefaultVal
        Exit Function
    End If
    Dim v As Long
    v = CLng(ValStr)
    If v >= MinVal And v <= MaxVal Then
        ParseLong = v
    Else
        ParseLong = DefaultVal
    End If
    Exit Function
Fallback:
    ParseLong = DefaultVal
End Function

Private Function ParseBool(ByVal ValStr As String, ByVal DefaultVal As Boolean) As Boolean
    Dim s As String
    s = UCase$(Trim$(ValStr))
    If s = "1" Or s = "TRUE" Then
        ParseBool = True
    ElseIf s = "0" Or s = "FALSE" Then
        ParseBool = False
    Else
        ParseBool = DefaultVal
    End If
End Function

' ==============================================================================
' Scalar Bridge Overloads for COM Automation Testing
' ==============================================================================

Public Sub SaveSettingsBridge( _
    ByVal ZeroStyle As Long, _
    ByVal ThousandStyle As Long, _
    ByVal FourStyle As Long, _
    ByVal DecMode As Long, _
    ByVal CurrType As Long, _
    ByVal AddChan As Boolean, _
    ByVal DecPlaces As Long, _
    ByVal Prefix As String, _
    ByVal Suffix As String, _
    ByVal SubUnit As String, _
    ByVal Casing As Long, _
    ByVal AddPeriod As Boolean, _
    ByVal OutputMode As Long, _
    ByVal QuickDir As Long, _
    ByVal ConfirmOverwrite As Boolean, _
    ByVal ShowSummary As Boolean, _
    ByVal SettingsVersion As Long _
)
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
    currOpts.CustomPrefix = Prefix
    currOpts.CustomSuffix = Suffix
    currOpts.CustomSubUnit = SubUnit
    
    fmtOpts.Casing = Casing
    fmtOpts.AddPeriod = AddPeriod
    
    appOpts.OutputMode = OutputMode
    appOpts.QuickConvertDirection = QuickDir
    appOpts.ConfirmOverwrite = ConfirmOverwrite
    appOpts.ShowBatchSummary = ShowSummary
    appOpts.SettingsVersion = SettingsVersion
    
    SaveAllSettings engOpts, currOpts, fmtOpts, appOpts
End Sub

Public Function LoadSettingsBridge() As Variant
    Dim engOpts As VnEngineOptions
    Dim currOpts As VnCurrencyOptions
    Dim fmtOpts As VnFormatOptions
    Dim appOpts As VnAppSettings
    
    LoadAllSettings engOpts, currOpts, fmtOpts, appOpts
    
    LoadSettingsBridge = Array( _
        CLng(engOpts.ZeroStyle), _
        CLng(engOpts.ThousandStyle), _
        CLng(engOpts.FourStyle), _
        CLng(engOpts.DecimalMode), _
        CLng(currOpts.CurrencyType), _
        CBool(currOpts.AddChan), _
        CLng(currOpts.DecimalPlaces), _
        CStr(currOpts.CustomPrefix), _
        CStr(currOpts.CustomSuffix), _
        CStr(currOpts.CustomSubUnit), _
        CLng(fmtOpts.Casing), _
        CBool(fmtOpts.AddPeriod), _
        CLng(appOpts.OutputMode), _
        CLng(appOpts.QuickConvertDirection), _
        CBool(appOpts.ConfirmOverwrite), _
        CBool(appOpts.ShowBatchSummary), _
        CLng(appOpts.SettingsVersion) _
    )
End Function

Public Sub ResetSettingsBridge()
    ResetSettingsToDefault
End Sub
