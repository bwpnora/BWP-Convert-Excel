Attribute VB_Name = "UDF"
Option Explicit

' ==============================================================================
' UDF.bas
' Worksheet User-Defined Functions (UDFs) for BWPConvertTTNVN:
' - BWPVNWORDS: General Vietnamese number-to-words reading with decimal digits.
' - BWPVND: Vietnamese Dong currency formatting (Sentence case, dong, optional chan).
' - BWPVNDUPPER: All-uppercase Vietnamese Dong currency formatting.
' - Convenience aliases: VNWORDS, VND, VNDUPPER.
'
' STRICT ARCHITECTURAL RULES:
' 1. Pure 7-bit ASCII safe source code (0 non-ASCII bytes).
' 2. Deterministic: UDFs must NEVER read Registry settings (Settings.bas).
'    Behavior is 100% determined by formula arguments alone.
' 3. Return type is Variant: returns CVErr(xlErrValue) on error / invalid input.
' 4. Blank/empty input cell returns "" (empty string).
' 5. Rejects arrays / multi-cell ranges in v1.0 (returns CVErr(xlErrValue)).
' 6. Application.Volatile is omitted (recalculates only when precedent cells change).
' ==============================================================================

' ------------------------------------------------------------------------------
' Canonical Worksheet Functions
' ------------------------------------------------------------------------------

Public Function BWPVNWORDS( _
    ByVal Target As Variant, _
    Optional ByVal ZeroStyle As Variant, _
    Optional ByVal ThousandStyle As Variant _
) As Variant
    BWPVNWORDS = ExecuteVnWords(Target, ZeroStyle, ThousandStyle)
End Function

Public Function BWPVND( _
    ByVal Target As Variant, _
    Optional ByVal AddChan As Variant = True, _
    Optional ByVal ZeroStyle As Variant, _
    Optional ByVal ThousandStyle As Variant _
) As Variant
    BWPVND = ExecuteVnD(Target, AddChan, ZeroStyle, ThousandStyle, VnCaseSentence)
End Function

Public Function BWPVNDUPPER( _
    ByVal Target As Variant, _
    Optional ByVal AddChan As Variant = True _
) As Variant
    BWPVNDUPPER = ExecuteVnD(Target, AddChan, Empty, Empty, VnCaseUpper)
End Function

' ------------------------------------------------------------------------------
' Convenience Aliases
' ------------------------------------------------------------------------------

Public Function VNWORDS( _
    ByVal Target As Variant, _
    Optional ByVal ZeroStyle As Variant, _
    Optional ByVal ThousandStyle As Variant _
) As Variant
    VNWORDS = ExecuteVnWords(Target, ZeroStyle, ThousandStyle)
End Function

Public Function VND( _
    ByVal Target As Variant, _
    Optional ByVal AddChan As Variant = True, _
    Optional ByVal ZeroStyle As Variant, _
    Optional ByVal ThousandStyle As Variant _
) As Variant
    VND = ExecuteVnD(Target, AddChan, ZeroStyle, ThousandStyle, VnCaseSentence)
End Function

Public Function VNDUPPER( _
    ByVal Target As Variant, _
    Optional ByVal AddChan As Variant = True _
) As Variant
    VNDUPPER = ExecuteVnD(Target, AddChan, Empty, Empty, VnCaseUpper)
End Function

' ------------------------------------------------------------------------------
' Private Execution & Argument Unwrapping Helpers
' ------------------------------------------------------------------------------

Private Function ExecuteVnWords( _
    ByVal Target As Variant, _
    ByVal ZeroStyle As Variant, _
    ByVal ThousandStyle As Variant _
) As Variant
    On Error GoTo ErrHandler

    Dim inputStatus As Long
    Dim numVal As Variant
    inputStatus = ValidateScalarInput(Target, numVal)

    If inputStatus = 1 Then
        ExecuteVnWords = ""
        Exit Function
    ElseIf inputStatus <> 0 Then
        ExecuteVnWords = CVErr(xlErrValue)
        Exit Function
    End If

    Dim optZero As Long
    If Not ParseStyleOption(ZeroStyle, VnZeroLe, optZero) Then
        ExecuteVnWords = CVErr(xlErrValue)
        Exit Function
    End If

    Dim optThousand As Long
    If Not ParseStyleOption(ThousandStyle, VnThousandNghin, optThousand) Then
        ExecuteVnWords = CVErr(xlErrValue)
        Exit Function
    End If

    Dim engOpts As VnEngineOptions
    engOpts = DefaultEngineOptions()
    engOpts.ZeroStyle = optZero
    engOpts.ThousandStyle = optThousand
    engOpts.DecimalMode = VnDecimalDigits

    Dim currOpts As VnCurrencyOptions
    currOpts = DefaultCurrencyOptions(VnCurrNone)
    currOpts.CurrencyType = VnCurrNone
    currOpts.DecimalPlaces = 0

    Dim fmtOpts As VnFormatOptions
    fmtOpts = DefaultFormatOptions()
    fmtOpts.Casing = VnCaseSentence
    fmtOpts.AddPeriod = True

    Dim outText As String
    Dim outErr As String
    If TryConvertNumber(numVal, engOpts, currOpts, fmtOpts, outText, outErr) Then
        ExecuteVnWords = outText
    Else
        ExecuteVnWords = CVErr(xlErrValue)
    End If
    Exit Function

ErrHandler:
    ExecuteVnWords = CVErr(xlErrValue)
End Function

Private Function ExecuteVnD( _
    ByVal Target As Variant, _
    ByVal AddChan As Variant, _
    ByVal ZeroStyle As Variant, _
    ByVal ThousandStyle As Variant, _
    ByVal Casing As Long _
) As Variant
    On Error GoTo ErrHandler

    Dim inputStatus As Long
    Dim rawNum As Variant
    inputStatus = ValidateScalarInput(Target, rawNum)

    If inputStatus = 1 Then
        ExecuteVnD = ""
        Exit Function
    ElseIf inputStatus <> 0 Then
        ExecuteVnD = CVErr(xlErrValue)
        Exit Function
    End If

    Dim bAddChan As Boolean
    If Not ParseAddChanOption(AddChan, bAddChan) Then
        ExecuteVnD = CVErr(xlErrValue)
        Exit Function
    End If

    Dim optZero As Long
    If Not ParseStyleOption(ZeroStyle, VnZeroLe, optZero) Then
        ExecuteVnD = CVErr(xlErrValue)
        Exit Function
    End If

    Dim optThousand As Long
    If Not ParseStyleOption(ThousandStyle, VnThousandNghin, optThousand) Then
        ExecuteVnD = CVErr(xlErrValue)
        Exit Function
    End If

    Dim numVal As Variant
    numVal = Fix(rawNum)

    Dim engOpts As VnEngineOptions
    engOpts = DefaultEngineOptions()
    engOpts.ZeroStyle = optZero
    engOpts.ThousandStyle = optThousand
    engOpts.DecimalMode = VnDecimalIgnore

    Dim currOpts As VnCurrencyOptions
    currOpts = DefaultCurrencyOptions(VnCurrVND)
    currOpts.CurrencyType = VnCurrVND
    currOpts.AddChan = bAddChan
    currOpts.DecimalPlaces = 0

    Dim fmtOpts As VnFormatOptions
    fmtOpts = DefaultFormatOptions()
    fmtOpts.Casing = Casing
    fmtOpts.AddPeriod = True

    Dim outText As String
    Dim outErr As String
    If TryConvertNumber(numVal, engOpts, currOpts, fmtOpts, outText, outErr) Then
        ExecuteVnD = outText
    Else
        ExecuteVnD = CVErr(xlErrValue)
    End If
    Exit Function

ErrHandler:
    ExecuteVnD = CVErr(xlErrValue)
End Function

Private Function UnwrapArg(ByVal ArgVal As Variant, ByRef OutVal As Variant) As Boolean
    On Error GoTo ErrHandler
    OutVal = Empty

    If IsObject(ArgVal) Then
        If ArgVal Is Nothing Then
            UnwrapArg = False
            Exit Function
        End If
        If TypeOf ArgVal Is Range Then
            If ArgVal.Areas.Count <> 1 Then
                UnwrapArg = False
                Exit Function
            End If
            If ArgVal.Rows.Count <> 1 Or ArgVal.Columns.Count <> 1 Then
                UnwrapArg = False
                Exit Function
            End If
            OutVal = ArgVal.Value2
            UnwrapArg = True
            Exit Function
        Else
            UnwrapArg = False
            Exit Function
        End If
    ElseIf IsArray(ArgVal) Then
        UnwrapArg = False
        Exit Function
    Else
        OutVal = ArgVal
        UnwrapArg = True
        Exit Function
    End If

ErrHandler:
    UnwrapArg = False
End Function

Private Function ValidateScalarInput(ByVal Target As Variant, ByRef OutVal As Variant) As Long
    ' Return codes:
    '  0 = Valid numeric value in OutVal
    '  1 = Blank/Empty (returns "")
    ' -1 = Invalid/Error (returns CVErr(xlErrValue))

    On Error GoTo ErrHandler
    OutVal = Empty

    Dim rawVal As Variant
    If Not UnwrapArg(Target, rawVal) Then
        ValidateScalarInput = -1
        Exit Function
    End If

    If IsError(rawVal) Then
        ValidateScalarInput = -1
        Exit Function
    End If

    If IsEmpty(rawVal) Or IsNull(rawVal) Then
        ValidateScalarInput = 1
        Exit Function
    End If

    If VarType(rawVal) = vbString Then
        If Len(Trim$(CStr(rawVal))) = 0 Then
            ValidateScalarInput = 1
            Exit Function
        End If
    End If

    Dim vt As Integer
    vt = VarType(rawVal)
    If vt = vbBoolean Or vt = vbDate Then
        ValidateScalarInput = -1
        Exit Function
    End If

    If Not IsNumeric(rawVal) Then
        ValidateScalarInput = -1
        Exit Function
    End If

    Dim dblCheck As Double
    dblCheck = CDbl(rawVal)
    If Abs(dblCheck) > MAX_SUPPORTED_VALUE Then
        ValidateScalarInput = -1
        Exit Function
    End If

    OutVal = rawVal
    ValidateScalarInput = 0
    Exit Function

ErrHandler:
    ValidateScalarInput = -1
End Function

Private Function ParseStyleOption(ByVal OptVal As Variant, ByVal DefaultVal As Long, ByRef OutVal As Long) As Boolean
    On Error GoTo ErrHandler

    If IsMissing(OptVal) Or IsEmpty(OptVal) Or IsNull(OptVal) Then
        OutVal = DefaultVal
        ParseStyleOption = True
        Exit Function
    End If

    Dim rawVal As Variant
    If Not UnwrapArg(OptVal, rawVal) Then
        ParseStyleOption = False
        Exit Function
    End If

    If IsMissing(rawVal) Or IsEmpty(rawVal) Or IsNull(rawVal) Then
        OutVal = DefaultVal
        ParseStyleOption = True
        Exit Function
    End If

    If Not IsNumeric(rawVal) Then
        ParseStyleOption = False
        Exit Function
    End If

    Dim v As Long
    v = CLng(rawVal)
    If v = 0 Or v = 1 Then
        OutVal = v
        ParseStyleOption = True
    Else
        ParseStyleOption = False
    End If
    Exit Function

ErrHandler:
    ParseStyleOption = False
End Function

Private Function ParseAddChanOption(ByVal ChanVal As Variant, ByRef OutChan As Boolean) As Boolean
    On Error GoTo ErrHandler

    If IsMissing(ChanVal) Or IsEmpty(ChanVal) Or IsNull(ChanVal) Then
        OutChan = True
        ParseAddChanOption = True
        Exit Function
    End If

    Dim rawVal As Variant
    If Not UnwrapArg(ChanVal, rawVal) Then
        ParseAddChanOption = False
        Exit Function
    End If

    If IsMissing(rawVal) Or IsEmpty(rawVal) Or IsNull(rawVal) Then
        OutChan = True
        ParseAddChanOption = True
        Exit Function
    End If

    If VarType(rawVal) = vbBoolean Then
        OutChan = CBool(rawVal)
        ParseAddChanOption = True
        Exit Function
    End If

    If IsNumeric(rawVal) Then
        OutChan = (CLng(rawVal) <> 0)
        ParseAddChanOption = True
        Exit Function
    End If

    If VarType(rawVal) = vbString Then
        Dim s As String
        s = UCase$(Trim$(CStr(rawVal)))
        If s = "TRUE" Or s = "1" Then
            OutChan = True
            ParseAddChanOption = True
            Exit Function
        ElseIf s = "FALSE" Or s = "0" Then
            OutChan = False
            ParseAddChanOption = True
            Exit Function
        Else
            ParseAddChanOption = False
            Exit Function
        End If
    End If

    ParseAddChanOption = False
    Exit Function

ErrHandler:
    ParseAddChanOption = False
End Function
