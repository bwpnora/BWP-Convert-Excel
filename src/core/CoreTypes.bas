Attribute VB_Name = "CoreTypes"
Option Explicit

' ==============================================================================
' CoreTypes.bas
' Global constants, error codes, enums, UDTs, and factory functions for BWPConvertTTNVN.
' Pure VBA: Zero Excel dependencies, zero external DLL dependencies.
' ==============================================================================

Public Const MAX_SUPPORTED_VALUE As Double = 999999999999999#

' Core Error Codes
Public Const ERR_INVALID_NUMBER As Long = vbObjectError + 2101
Public Const ERR_OUT_OF_RANGE As Long = vbObjectError + 2102
Public Const ERR_INVALID_OPTIONS As Long = vbObjectError + 2103

' Enums
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

' User-Defined Types (UDTs)
Public Type VnEngineOptions
    ZeroStyle As VnZeroStyle
    ThousandStyle As VnThousandStyle
    FourStyle As VnFourStyle
    DecimalMode As VnDecimalMode
End Type

Public Type VnCurrencyOptions
    CurrencyType As VnCurrencyType
    AddChan As Boolean
    DecimalPlaces As Integer
    CustomPrefix As String
    CustomSuffix As String
    CustomSubUnit As String
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
            Opts.DecimalPlaces = 0
        Case VnCurrUSD
            Opts.DecimalPlaces = 2
        Case Else
            Opts.DecimalPlaces = 0
    End Select
    DefaultCurrencyOptions = Opts
End Function

Public Function DefaultFormatOptions() As VnFormatOptions
    Dim Opts As VnFormatOptions
    Opts.Casing = VnCaseSentence
    Opts.AddPeriod = True
    DefaultFormatOptions = Opts
End Function
