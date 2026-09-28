Attribute VB_Name = "CoreCoordinator"
Option Explicit

' ==============================================================================
' CoreCoordinator.bas
' Top-level unified coordinator for BWPConvertTTNVN conversion pipeline:
' Value -> Validate -> Currency/Number Words -> TextFormatter -> FinalText
' Pure VBA: Zero Excel dependencies, zero UI, zero external DLL dependencies.
' All non-ASCII characters accessed programmatically via UnicodeText.bas.
' Source code is 100% 7-bit ASCII safe.
' ==============================================================================

Public Function ConvertNumber( _
    ByVal Value As Variant, _
    ByRef EngineOpts As VnEngineOptions, _
    ByRef CurrOpts As VnCurrencyOptions, _
    ByRef FormatOpts As VnFormatOptions, _
    ByRef OutErrorMessage As String _
) As String
    Dim outText As String
    If TryConvertNumber(Value, EngineOpts, CurrOpts, FormatOpts, outText, OutErrorMessage) Then
        ConvertNumber = outText
    Else
        ConvertNumber = ""
    End If
End Function

Public Function TryConvertNumber( _
    ByVal Value As Variant, _
    ByRef EngineOpts As VnEngineOptions, _
    ByRef CurrOpts As VnCurrencyOptions, _
    ByRef FormatOpts As VnFormatOptions, _
    ByRef OutText As String, _
    ByRef OutErrorMessage As String _
) As Boolean
    On Error GoTo ErrHandler
    OutText = ""
    OutErrorMessage = ""
    
    ' 1. Convert to currency / number words
    Dim rawWords As String
    rawWords = NumberToCurrencyWords(Value, CurrOpts, EngineOpts)
    
    ' 2. Format text (whitespace, casing, punctuation)
    OutText = FormatText(rawWords, FormatOpts.Casing, FormatOpts.AddPeriod)
    TryConvertNumber = True
    Exit Function

ErrHandler:
    OutText = ""
    OutErrorMessage = Err.Description
    If Len(OutErrorMessage) = 0 Then
        OutErrorMessage = "Loi chuyen doi: " & CStr(Err.Number)
    End If
    TryConvertNumber = False
End Function

Public Function ConvertNumberDefault(ByVal Value As Variant) As String
    Dim engOpts As VnEngineOptions
    Dim currOpts As VnCurrencyOptions
    Dim fmtOpts As VnFormatOptions
    Dim errMsg As String
    
    engOpts = DefaultEngineOptions()
    currOpts = DefaultCurrencyOptions(VnCurrVND)
    fmtOpts = DefaultFormatOptions()
    
    ConvertNumberDefault = ConvertNumber(Value, engOpts, currOpts, fmtOpts, errMsg)
End Function

' ------------------------------------------------------------------------------
' Scalar Bridge Overloads for COM Automation Testing
' ------------------------------------------------------------------------------

Public Function ConvertNumberWithOptions( _
    ByVal Value As Variant, _
    ByVal CurrType As Long, _
    ByVal AddChan As Boolean, _
    ByVal DecPlaces As Long, _
    ByVal ZeroStyle As Long, _
    ByVal ThousandStyle As Long, _
    ByVal FourStyle As Long, _
    ByVal DecMode As Long, _
    ByVal Casing As Long, _
    ByVal AddPeriod As Boolean _
) As String
    Dim engOpts As VnEngineOptions
    Dim currOpts As VnCurrencyOptions
    Dim fmtOpts As VnFormatOptions
    Dim errMsg As String
    
    engOpts.ZeroStyle = ZeroStyle
    engOpts.ThousandStyle = ThousandStyle
    engOpts.FourStyle = FourStyle
    engOpts.DecimalMode = DecMode
    
    currOpts.CurrencyType = CurrType
    currOpts.AddChan = AddChan
    currOpts.DecimalPlaces = DecPlaces
    currOpts.CustomPrefix = ""
    currOpts.CustomSuffix = ""
    currOpts.CustomSubUnit = ""
    
    fmtOpts.Casing = Casing
    fmtOpts.AddPeriod = AddPeriod
    
    ConvertNumberWithOptions = ConvertNumber(Value, engOpts, currOpts, fmtOpts, errMsg)
End Function

Public Function TryConvertNumberWithOptions( _
    ByVal Value As Variant, _
    ByVal CurrType As Long, _
    ByVal AddChan As Boolean, _
    ByVal DecPlaces As Long, _
    ByVal ZeroStyle As Long, _
    ByVal ThousandStyle As Long, _
    ByVal FourStyle As Long, _
    ByVal DecMode As Long, _
    ByVal Casing As Long, _
    ByVal AddPeriod As Boolean, _
    ByRef OutErrorMessage As String _
) As String
    Dim engOpts As VnEngineOptions
    Dim currOpts As VnCurrencyOptions
    Dim fmtOpts As VnFormatOptions
    Dim outText As String
    
    engOpts.ZeroStyle = ZeroStyle
    engOpts.ThousandStyle = ThousandStyle
    engOpts.FourStyle = FourStyle
    engOpts.DecimalMode = DecMode
    
    currOpts.CurrencyType = CurrType
    currOpts.AddChan = AddChan
    currOpts.DecimalPlaces = DecPlaces
    currOpts.CustomPrefix = ""
    currOpts.CustomSuffix = ""
    currOpts.CustomSubUnit = ""
    
    fmtOpts.Casing = Casing
    fmtOpts.AddPeriod = AddPeriod
    
    If TryConvertNumber(Value, engOpts, currOpts, fmtOpts, outText, OutErrorMessage) Then
        TryConvertNumberWithOptions = outText
    Else
        TryConvertNumberWithOptions = ""
    End If
End Function
