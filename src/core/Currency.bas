Attribute VB_Name = "VnCurrency"
Option Explicit

' ==============================================================================
' Currency.bas
' Pure VBA currency formatting, deterministic rounding, sub-units, and "chan".
' Zero Excel dependencies, zero UI, zero external DLL dependencies.
' All non-ASCII characters accessed programmatically via UnicodeText.bas.
' Source code is 100% 7-bit ASCII safe.
' ==============================================================================

' ------------------------------------------------------------------------------
' Public API: Deterministic Rounding
' ------------------------------------------------------------------------------

Public Function RoundHalfAwayFromZero(ByVal Value As Variant, ByVal DecimalPlaces As Long) As Variant
    If IsEmpty(Value) Or IsNull(Value) Or IsError(Value) Or IsObject(Value) Then
        Err.Raise ERR_INVALID_NUMBER, "Currency.RoundHalfAwayFromZero", "Gia tri khong hop le."
    End If
    
    Dim vt As Integer
    vt = VarType(Value)
    If vt = vbBoolean Or vt = vbDate Or Not IsNumeric(Value) Then
        Err.Raise ERR_INVALID_NUMBER, "Currency.RoundHalfAwayFromZero", "Gia tri khong phai la so hop le."
    End If
    
    Dim dblCheck As Double
    On Error GoTo OverflowErr
    dblCheck = CDbl(Value)
    On Error GoTo 0
    If Abs(dblCheck) > MAX_SUPPORTED_VALUE Then
        Err.Raise ERR_OUT_OF_RANGE, "Currency.RoundHalfAwayFromZero", "Gia tri vuot qua gioi han ho tro (toi da 999 nghin ty)."
    End If
    
    If DecimalPlaces < 0 Then
        DecimalPlaces = 0
    End If
    
    Dim dVal As Variant
    dVal = CDec(Value)
    
    Dim factor As Variant
    factor = CDec(10 ^ DecimalPlaces)
    
    Dim scaled As Variant
    scaled = dVal * factor
    
    Dim roundedScaled As Variant
    If scaled >= 0 Then
        roundedScaled = Fix(scaled + CDec(0.5))
    Else
        roundedScaled = Fix(scaled - CDec(0.5))
    End If
    
    Dim result As Variant
    If DecimalPlaces = 0 Then
        result = roundedScaled
    Else
        result = roundedScaled / factor
    End If
    
    ' Normalize -0 to 0
    If result = 0 Then
        RoundHalfAwayFromZero = CDec(0)
    Else
        RoundHalfAwayFromZero = result
    End If
    Exit Function

OverflowErr:
    Err.Raise ERR_OUT_OF_RANGE, "Currency.RoundHalfAwayFromZero", "Gia tri vuot qua gioi han so."
End Function

' ------------------------------------------------------------------------------
' Public API: Currency to Words
' ------------------------------------------------------------------------------

Public Function NumberToCurrencyWords( _
    ByVal Value As Variant, _
    ByRef CurrOptions As VnCurrencyOptions, _
    ByRef EngineOptions As VnEngineOptions _
) As String
    ' 1. Type and bounds validation
    Dim errMsg As String
    If Not IsValidNumber(Value, errMsg) Then
        If InStr(errMsg, "vuot qua gioi han") > 0 Then
            Err.Raise ERR_OUT_OF_RANGE, "Currency.NumberToCurrencyWords", errMsg
        Else
            Err.Raise ERR_INVALID_NUMBER, "Currency.NumberToCurrencyWords", errMsg
        End If
    End If
    
    ' 2. Options validation
    If CurrOptions.CurrencyType < 0 Or CurrOptions.CurrencyType > 3 Then
        Err.Raise ERR_INVALID_OPTIONS, "Currency.NumberToCurrencyWords", "Loai tien te khong hop le."
    End If
    If CurrOptions.DecimalPlaces < 0 Or CurrOptions.DecimalPlaces > 28 Then
        Err.Raise ERR_INVALID_OPTIONS, "Currency.NumberToCurrencyWords", "So chu so thap phan khong hop le."
    End If
    If EngineOptions.ZeroStyle < 0 Or EngineOptions.ZeroStyle > 1 Or _
       EngineOptions.ThousandStyle < 0 Or EngineOptions.ThousandStyle > 1 Or _
       EngineOptions.FourStyle < 0 Or EngineOptions.FourStyle > 1 Or _
       EngineOptions.DecimalMode < 0 Or EngineOptions.DecimalMode > 1 Then
        Err.Raise ERR_INVALID_OPTIONS, "Currency.NumberToCurrencyWords", "Tuy chon engine khong hop le."
    End If
    
    ' 3. VnCurrNone: reads base number words
    If CurrOptions.CurrencyType = VnCurrNone Then
        If CurrOptions.DecimalPlaces > 0 And EngineOptions.DecimalMode = VnDecimalDigits Then
            Dim roundedBase As Variant
            roundedBase = RoundHalfAwayFromZero(Value, CurrOptions.DecimalPlaces)
            NumberToCurrencyWords = NumberToVietnamese(roundedBase, EngineOptions)
        Else
            NumberToCurrencyWords = NumberToVietnamese(Value, EngineOptions)
        End If
        Exit Function
    End If
    
    ' 4. Determine decimal places for rounding
    Dim decPlaces As Long
    decPlaces = CurrOptions.DecimalPlaces
    If decPlaces < 0 Then decPlaces = 0
    
    ' 5. Deterministic carry rounding before decomposition
    Dim roundedVal As Variant
    roundedVal = RoundHalfAwayFromZero(Value, decPlaces)
    
    Dim dblRoundCheck As Double
    On Error GoTo OverflowErr
    dblRoundCheck = CDbl(roundedVal)
    On Error GoTo 0
    If Abs(dblRoundCheck) > MAX_SUPPORTED_VALUE Then
        Err.Raise ERR_OUT_OF_RANGE, "Currency.NumberToCurrencyWords", "Gia tri vuot qua gioi han ho tro (toi da 999 nghin ty)."
    End If
    
    ' 6. Inspect sign once
    Dim isNegative As Boolean
    isNegative = (roundedVal < 0)
    
    Dim absVal As Variant
    absVal = Abs(CDec(roundedVal))
    
    ' 7. Decompose into integer part and sub-units
    Dim intVal As Variant
    intVal = Fix(absVal)
    
    Dim fracVal As Variant
    fracVal = absVal - intVal
    
    Dim subUnitVal As Long
    subUnitVal = 0
    If decPlaces > 0 Then
        subUnitVal = CLng(RoundHalfAwayFromZero(fracVal * CDec(10 ^ decPlaces), 0))
    End If
    
    ' Read integer part words with DecimalMode = VnDecimalIgnore
    Dim engOptsInt As VnEngineOptions
    engOptsInt = EngineOptions
    engOptsInt.DecimalMode = VnDecimalIgnore
    
    Dim intWords As String
    intWords = NumberToVietnamese(intVal, engOptsInt)
    
    ' 8. Assemble currency phrase
    Dim result As String
    Select Case CurrOptions.CurrencyType
        Case VnCurrVND
            result = intWords & " " & WordDong()
            If subUnitVal > 0 Then
                Dim vndSubWords As String
                vndSubWords = NumberToVietnamese(subUnitVal, engOptsInt)
                result = result & " " & vndSubWords & " " & WordXu()
            Else
                If CurrOptions.AddChan Then
                    result = result & " " & WordChan()
                End If
            End If
            
        Case VnCurrUSD
            result = intWords & " " & WordDolaMy()
            If subUnitVal > 0 Then
                Dim usdSubWords As String
                usdSubWords = NumberToVietnamese(subUnitVal, engOptsInt)
                result = result & " " & usdSubWords & " " & WordCent()
            Else
                If CurrOptions.AddChan Then
                    result = result & " " & WordChan()
                End If
            End If
            
        Case VnCurrCustom
            result = intWords
            If Len(CurrOptions.CustomSuffix) > 0 Then
                result = result & " " & CurrOptions.CustomSuffix
            End If
            
            If subUnitVal > 0 Then
                Dim customSubWords As String
                customSubWords = NumberToVietnamese(subUnitVal, engOptsInt)
                result = result & " " & customSubWords
                If Len(CurrOptions.CustomSubUnit) > 0 Then
                    result = result & " " & CurrOptions.CustomSubUnit
                End If
            Else
                If CurrOptions.AddChan Then
                    result = result & " " & WordChan()
                End If
            End If
    End Select
    
    ' 9. Prefix negative sign before entire currency phrase
    If isNegative Then
        result = ToUnicodeLower(WordAm()) & " " & result
    End If
    
    ' 10. Prefix custom prefix if defined
    If CurrOptions.CurrencyType = VnCurrCustom Then
        If Len(CurrOptions.CustomPrefix) > 0 Then
            result = CurrOptions.CustomPrefix & " " & result
        End If
    End If
    
    NumberToCurrencyWords = result
    Exit Function

OverflowErr:
    Err.Raise ERR_OUT_OF_RANGE, "Currency.NumberToCurrencyWords", "Gia tri vuot qua gioi han so."
End Function

' ------------------------------------------------------------------------------
' Scalar Bridge Overload for COM Automation Testing
' ------------------------------------------------------------------------------

Public Function NumberToCurrencyWordsEx( _
    ByVal Value As Variant, _
    ByVal CurrType As Long, _
    ByVal AddChan As Boolean, _
    ByVal DecPlaces As Long, _
    ByVal CustomPrefix As String, _
    ByVal CustomSuffix As String, _
    ByVal CustomSubUnit As String, _
    ByVal ZeroStyle As Long, _
    ByVal ThousandStyle As Long, _
    ByVal FourStyle As Long, _
    ByVal DecMode As Long _
) As String
    Dim currOpts As VnCurrencyOptions
    Dim engOpts As VnEngineOptions
    
    currOpts.CurrencyType = CurrType
    currOpts.AddChan = AddChan
    currOpts.DecimalPlaces = DecPlaces
    currOpts.CustomPrefix = CustomPrefix
    currOpts.CustomSuffix = CustomSuffix
    currOpts.CustomSubUnit = CustomSubUnit
    
    engOpts.ZeroStyle = ZeroStyle
    engOpts.ThousandStyle = ThousandStyle
    engOpts.FourStyle = FourStyle
    engOpts.DecimalMode = DecMode
    
    NumberToCurrencyWordsEx = NumberToCurrencyWords(Value, currOpts, engOpts)
End Function
