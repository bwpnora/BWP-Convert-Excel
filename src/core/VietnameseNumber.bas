Attribute VB_Name = "VietnameseNumber"
Option Explicit

' ==============================================================================
' VietnameseNumber.bas
' Pure VBA algorithmic Vietnamese number-to-words conversion engine.
' Zero Excel dependencies, zero UI, zero currency concepts, zero external DLLs.
' All non-ASCII characters accessed programmatically via UnicodeText.bas.
' Source code is 100% 7-bit ASCII safe.
' ==============================================================================

' ------------------------------------------------------------------------------
' Public API: Validation
' ------------------------------------------------------------------------------

Public Function IsValidNumber(ByVal Value As Variant, ByRef OutErrorMessage As String) As Boolean
    OutErrorMessage = ""
    
    If IsEmpty(Value) Then
        OutErrorMessage = "Gia tri khong duoc de trong."
        IsValidNumber = False
        Exit Function
    End If
    
    If IsNull(Value) Then
        OutErrorMessage = "Gia tri khong duoc Null."
        IsValidNumber = False
        Exit Function
    End If
    
    If IsError(Value) Then
        OutErrorMessage = "Gia tri bi loi."
        IsValidNumber = False
        Exit Function
    End If
    
    If IsObject(Value) Then
        OutErrorMessage = "Gia tri khong hop le (Object)."
        IsValidNumber = False
        Exit Function
    End If
    
    Dim vt As Integer
    vt = VarType(Value)
    
    If vt = vbBoolean Then
        OutErrorMessage = "Gia tri kieu Boolean khong duoc ho tro."
        IsValidNumber = False
        Exit Function
    End If
    
    If vt = vbDate Then
        OutErrorMessage = "Gia tri kieu Date khong duoc ho tro."
        IsValidNumber = False
        Exit Function
    End If
    
    If Not IsNumeric(Value) Then
        OutErrorMessage = "Gia tri khong phai la so hop le."
        IsValidNumber = False
        Exit Function
    End If
    
    Dim dblVal As Double
    On Error GoTo RangeErr
    dblVal = CDbl(Value)
    On Error GoTo 0
    
    If Abs(dblVal) > MAX_SUPPORTED_VALUE Then
        OutErrorMessage = "Gia tri vuot qua gioi han ho tro (toi da 999 nghin ty)."
        IsValidNumber = False
        Exit Function
    End If
    
    IsValidNumber = True
    Exit Function

RangeErr:
    OutErrorMessage = "Gia tri vuot qua gioi han so."
    IsValidNumber = False
End Function

' ------------------------------------------------------------------------------
' Public API: Core Conversions
' ------------------------------------------------------------------------------

Public Function NumberToVietnamese(ByVal Value As Variant, ByRef Options As VnEngineOptions) As String
    ' 1. Type and format validation
    If IsEmpty(Value) Or IsNull(Value) Or IsError(Value) Or IsObject(Value) Then
        Err.Raise ERR_INVALID_NUMBER, "VietnameseNumber.NumberToVietnamese", "Gia tri khong phai la so hop le."
    End If
    
    Dim vt As Integer
    vt = VarType(Value)
    If vt = vbBoolean Or vt = vbDate Or Not IsNumeric(Value) Then
        Err.Raise ERR_INVALID_NUMBER, "VietnameseNumber.NumberToVietnamese", "Gia tri khong phai la so hop le."
    End If
    
    Dim dblVal As Double
    On Error GoTo OverflowErr
    dblVal = CDbl(Value)
    On Error GoTo 0
    
    If Abs(dblVal) > MAX_SUPPORTED_VALUE Then
        Err.Raise ERR_OUT_OF_RANGE, "VietnameseNumber.NumberToVietnamese", "Gia tri vuot qua gioi han ho tro (toi da 999 nghin ty)."
    End If
    
    ' 2. Options validation
    If Options.ZeroStyle < 0 Or Options.ZeroStyle > 1 Or _
       Options.ThousandStyle < 0 Or Options.ThousandStyle > 1 Or _
       Options.FourStyle < 0 Or Options.FourStyle > 1 Or _
       Options.DecimalMode < 0 Or Options.DecimalMode > 1 Then
        Err.Raise ERR_INVALID_OPTIONS, "VietnameseNumber.NumberToVietnamese", "Tuy chon engine khong hop le."
    End If
    
    ' 3. Sign handling
    Dim isNegative As Boolean
    isNegative = (dblVal < 0#)
    
    Dim absVal As Double
    absVal = Abs(dblVal)
    
    ' 4. Integer and decimal decomposition
    Dim intVal As Double
    intVal = Fix(absVal)
    
    Dim intWords As String
    intWords = ConvertIntegerPart(intVal, Options)
    
    ' 5. Decimal digits processing
    Dim result As String
    If Options.DecimalMode = VnDecimalDigits Then
        Dim decStr As String
        decStr = ExtractDecimalDigits(Value, absVal)
        
        If Len(decStr) > 0 Then
            Dim decWords As String
            decWords = ReadDecimalDigits(decStr)
            If Len(decWords) > 0 Then
                result = intWords & " " & WordPhay() & " " & decWords
            Else
                result = intWords
            End If
        Else
            result = intWords
        End If
    Else
        result = intWords
    End If
    
    ' 6. Apply negative prefix
    If isNegative Then
        If result <> WordKhong() Or Options.DecimalMode = VnDecimalDigits Then
            result = ToUnicodeLower(WordAm()) & " " & result
        End If
    End If
    
    NumberToVietnamese = result
    Exit Function

OverflowErr:
    Err.Raise ERR_OUT_OF_RANGE, "VietnameseNumber.NumberToVietnamese", "Gia tri vuot qua gioi han so."
End Function

Public Function TryNumberToVietnamese( _
    ByVal Value As Variant, _
    ByRef Options As VnEngineOptions, _
    ByRef OutText As String, _
    ByRef OutErrorMessage As String _
) As Boolean
    On Error GoTo ErrHandler
    OutText = ""
    OutErrorMessage = ""
    
    If Not IsValidNumber(Value, OutErrorMessage) Then
        TryNumberToVietnamese = False
        Exit Function
    End If
    
    OutText = NumberToVietnamese(Value, Options)
    TryNumberToVietnamese = True
    Exit Function

ErrHandler:
    OutText = ""
    OutErrorMessage = Err.Description
    If Len(OutErrorMessage) = 0 Then
        OutErrorMessage = "Loi chuyen doi so: " & CStr(Err.Number)
    End If
    TryNumberToVietnamese = False
End Function

Public Function NumberToVietnameseDefault(ByVal Value As Variant) As String
    Dim Opts As VnEngineOptions
    Opts = DefaultEngineOptions()
    NumberToVietnameseDefault = NumberToVietnamese(Value, Opts)
End Function

Public Function NumberToVietnameseEx( _
    ByVal Value As Variant, _
    Optional ByVal ZeroStyle As Long = 0, _
    Optional ByVal ThousandStyle As Long = 0, _
    Optional ByVal FourStyle As Long = 0, _
    Optional ByVal DecimalMode As Long = 0 _
) As String
    Dim Opts As VnEngineOptions
    Opts.ZeroStyle = ZeroStyle
    Opts.ThousandStyle = ThousandStyle
    Opts.FourStyle = FourStyle
    Opts.DecimalMode = DecimalMode
    NumberToVietnameseEx = NumberToVietnamese(Value, Opts)
End Function

Public Function TryNumberToVietnameseEx( _
    ByVal Value As Variant, _
    ByVal ZeroStyle As Long, _
    ByVal ThousandStyle As Long, _
    ByVal FourStyle As Long, _
    ByVal DecimalMode As Long, _
    ByRef OutErrorMessage As String _
) As String
    Dim Opts As VnEngineOptions
    Dim OutText As String
    Opts.ZeroStyle = ZeroStyle
    Opts.ThousandStyle = ThousandStyle
    Opts.FourStyle = FourStyle
    Opts.DecimalMode = DecimalMode
    If TryNumberToVietnamese(Value, Opts, OutText, OutErrorMessage) Then
        TryNumberToVietnameseEx = OutText
    Else
        TryNumberToVietnameseEx = ""
    End If
End Function

' ------------------------------------------------------------------------------
' Private Helper Functions
' ------------------------------------------------------------------------------

Private Function DigitWord(ByVal d As Integer) As String
    Select Case d
        Case 0: DigitWord = WordKhong()
        Case 1: DigitWord = WordMot()
        Case 2: DigitWord = WordHai()
        Case 3: DigitWord = WordBa()
        Case 4: DigitWord = WordBon()
        Case 5: DigitWord = WordNam()
        Case 6: DigitWord = WordSau()
        Case 7: DigitWord = WordBay()
        Case 8: DigitWord = WordTam()
        Case 9: DigitWord = WordChin()
        Case Else: DigitWord = ""
    End Select
End Function

Private Function GetScaleWord(ByVal GroupIndex As Integer, ByRef Options As VnEngineOptions) As String
    Select Case GroupIndex
        Case 0
            GetScaleWord = ""
        Case 1
            If Options.ThousandStyle = VnThousandNgan Then
                GetScaleWord = WordNgan()
            Else
                GetScaleWord = WordNghin()
            End If
        Case 2
            GetScaleWord = WordTrieu()
        Case 3
            GetScaleWord = WordTy()
        Case 4
            If Options.ThousandStyle = VnThousandNgan Then
                GetScaleWord = WordNgan() & " " & WordTy()
            Else
                GetScaleWord = WordNghin() & " " & WordTy()
            End If
        Case Else
            GetScaleWord = ""
    End Select
End Function

Private Function ReadTriplet( _
    ByVal h As Integer, _
    ByVal t As Integer, _
    ByVal u As Integer, _
    ByVal forceFullTriplet As Boolean, _
    ByRef Options As VnEngineOptions _
) As String
    If h = 0 And t = 0 And u = 0 Then
        ReadTriplet = ""
        Exit Function
    End If

    Dim hPart As String
    If h > 0 Then
        hPart = DigitWord(h) & " " & WordTram()
    ElseIf forceFullTriplet Then
        hPart = WordKhong() & " " & WordTram()
    Else
        hPart = ""
    End If

    Dim tuPart As String
    If t = 0 And u = 0 Then
        tuPart = ""
    ElseIf t = 0 Then
        ' Tens = 0, Units > 0
        Dim zeroWord As String
        If Options.ZeroStyle = VnZeroLinh Then
            zeroWord = WordLinh()
        Else
            zeroWord = WordLe()
        End If
        
        Dim uWord As String
        Select Case u
            Case 1: uWord = WordMot()
            Case 4: uWord = WordBon()
            Case 5: uWord = WordNam()
            Case Else: uWord = DigitWord(u)
        End Select
        
        If Len(hPart) > 0 Then
            tuPart = zeroWord & " " & uWord
        Else
            tuPart = uWord
        End If
    ElseIf t = 1 Then
        ' Tens = 1 (muoi)
        Dim tWord As String
        tWord = WordMuoi()
        
        Select Case u
            Case 0: tuPart = tWord
            Case 1: tuPart = tWord & " " & WordMot()
            Case 5: tuPart = tWord & " " & WordLam()
            Case Else: tuPart = tWord & " " & DigitWord(u)
        End Select
    Else
        ' Tens >= 2 (hai muoi, ba muoi...)
        Dim tensWord As String
        tensWord = DigitWord(t) & " " & WordMuoiChuc()
        
        Select Case u
            Case 0
                tuPart = tensWord
            Case 1
                tuPart = tensWord & " " & WordMotCuoi()
            Case 4
                If Options.FourStyle = VnFourBon Then
                    tuPart = tensWord & " " & WordBon()
                Else
                    tuPart = tensWord & " " & WordTu()
                End If
            Case 5
                tuPart = tensWord & " " & WordLam()
            Case Else
                tuPart = tensWord & " " & DigitWord(u)
        End Select
    End If

    If Len(hPart) > 0 And Len(tuPart) > 0 Then
        ReadTriplet = hPart & " " & tuPart
    ElseIf Len(hPart) > 0 Then
        ReadTriplet = hPart
    Else
        ReadTriplet = tuPart
    End If
End Function

Private Function ConvertIntegerPart(ByVal absVal As Double, ByRef Options As VnEngineOptions) As String
    If absVal = 0# Then
        ConvertIntegerPart = WordKhong()
        Exit Function
    End If

    Dim intStr As String
    intStr = Format$(absVal, "0")
    
    Dim remLen As Integer
    remLen = Len(intStr) Mod 3
    If remLen = 1 Then
        intStr = "00" & intStr
    ElseIf remLen = 2 Then
        intStr = "0" & intStr
    End If
    
    Dim totalGroups As Integer
    totalGroups = Len(intStr) \ 3
    
    Dim result As String
    result = ""
    
    Dim hasHigherNonZero As Boolean
    hasHigherNonZero = False
    
    Dim k As Integer
    Dim pos As Integer
    Dim h As Integer, t As Integer, u As Integer
    Dim tripStr As String
    Dim tripWords As String
    Dim scaleWord As String
    
    For k = totalGroups - 1 To 0 Step -1
        pos = (totalGroups - 1 - k) * 3 + 1
        tripStr = Mid$(intStr, pos, 3)
        h = CInt(Mid$(tripStr, 1, 1))
        t = CInt(Mid$(tripStr, 2, 1))
        u = CInt(Mid$(tripStr, 3, 1))
        
        If h = 0 And t = 0 And u = 0 Then
            ' Triplet 000 is silent.
            ' hasHigherNonZero remains unchanged.
        Else
            tripWords = ReadTriplet(h, t, u, hasHigherNonZero, Options)
            scaleWord = GetScaleWord(k, Options)
            
            If Len(scaleWord) > 0 Then
                tripWords = tripWords & " " & scaleWord
            End If
            
            If Len(result) > 0 Then
                result = result & " " & tripWords
            Else
                result = tripWords
            End If
            
            hasHigherNonZero = True
        End If
    Next k
    
    ConvertIntegerPart = result
End Function

Private Function ExtractDecimalDigits(ByVal RawValue As Variant, ByVal AbsVal As Double) As String
    Dim s As String
    If VarType(RawValue) = vbString Then
        If InStr(CStr(RawValue), ".") > 0 Then
            s = CStr(RawValue)
        Else
            s = Trim$(Str$(AbsVal))
        End If
    Else
        s = Trim$(Str$(AbsVal))
    End If
    
    Dim dotPos As Long
    dotPos = InStr(s, ".")
    If dotPos > 0 Then
        Dim rawDec As String
        rawDec = Mid$(s, dotPos + 1)
        
        ' Filter strictly digits
        Dim cleanDec As String
        cleanDec = ""
        Dim i As Long
        Dim ch As String
        For i = 1 To Len(rawDec)
            ch = Mid$(rawDec, i, 1)
            If ch >= "0" And ch <= "9" Then
                cleanDec = cleanDec & ch
            Else
                Exit For
            End If
        Next i
        ExtractDecimalDigits = cleanDec
    Else
        ExtractDecimalDigits = ""
    End If
End Function

Private Function ReadDecimalDigits(ByVal decStr As String) As String
    Dim result As String
    result = ""
    Dim i As Integer
    Dim ch As String
    Dim w As String
    
    For i = 1 To Len(decStr)
        ch = Mid$(decStr, i, 1)
        Select Case ch
            Case "0": w = WordKhong()
            Case "1": w = WordMot()
            Case "2": w = WordHai()
            Case "3": w = WordBa()
            Case "4": w = WordBon()
            Case "5": w = WordNam()
            Case "6": w = WordSau()
            Case "7": w = WordBay()
            Case "8": w = WordTam()
            Case "9": w = WordChin()
            Case Else: w = ""
        End Select
        If Len(w) > 0 Then
            If Len(result) > 0 Then
                result = result & " " & w
            Else
                result = w
            End If
        End If
    Next i
    
    ReadDecimalDigits = result
End Function
