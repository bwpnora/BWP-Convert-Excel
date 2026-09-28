Attribute VB_Name = "test_engine"
Option Explicit

' ==============================================================================
' test_engine.bas
' Pure VBA automated test runner for VietnameseNumber.bas.
' Zero Excel dependencies, zero UI, zero external DLL dependencies.
' All non-ASCII characters accessed via UnicodeText.bas.
' Source code is 100% 7-bit ASCII safe.
' ==============================================================================

Private m_PassedCount As Long
Private m_FailedCount As Long
Private m_FirstFailure As String

' ------------------------------------------------------------------------------
' Bridge functions for external automation (Excel.Run)
' ------------------------------------------------------------------------------

Public Function NumberToVietnameseWithOptions( _
    ByVal Value As Variant, _
    ByVal ZeroStyle As Long, _
    ByVal ThousandStyle As Long, _
    ByVal FourStyle As Long, _
    ByVal DecimalMode As Long _
) As String
    Dim Opts As VnEngineOptions
    Opts.ZeroStyle = ZeroStyle
    Opts.ThousandStyle = ThousandStyle
    Opts.FourStyle = FourStyle
    Opts.DecimalMode = DecimalMode
    NumberToVietnameseWithOptions = NumberToVietnamese(Value, Opts)
End Function

Public Function TryNumberToVietnameseWithOptions( _
    ByVal Value As Variant, _
    ByVal ZeroStyle As Long, _
    ByVal ThousandStyle As Long, _
    ByVal FourStyle As Long, _
    ByVal DecimalMode As Long _
) As String
    Dim Opts As VnEngineOptions
    Dim OutText As String
    Dim OutErrMsg As String
    Opts.ZeroStyle = ZeroStyle
    Opts.ThousandStyle = ThousandStyle
    Opts.FourStyle = FourStyle
    Opts.DecimalMode = DecimalMode
    If TryNumberToVietnamese(Value, Opts, OutText, OutErrMsg) Then
        TryNumberToVietnameseWithOptions = OutText
    Else
        TryNumberToVietnameseWithOptions = ""
    End If
End Function

Public Function TryConvertGetError( _
    ByVal Value As Variant, _
    ByVal ZeroStyle As Long, _
    ByVal ThousandStyle As Long, _
    ByVal FourStyle As Long, _
    ByVal DecimalMode As Long _
) As String
    Dim Opts As VnEngineOptions
    Dim OutText As String
    Dim OutErrMsg As String
    Opts.ZeroStyle = ZeroStyle
    Opts.ThousandStyle = ThousandStyle
    Opts.FourStyle = FourStyle
    Opts.DecimalMode = DecimalMode
    If TryNumberToVietnamese(Value, Opts, OutText, OutErrMsg) Then
        TryConvertGetError = ""
    Else
        TryConvertGetError = OutErrMsg
    End If
End Function

Public Function GetNumberToVietnameseErrorCode( _
    ByVal Value As Variant, _
    ByVal ZeroStyle As Long, _
    ByVal ThousandStyle As Long, _
    ByVal FourStyle As Long, _
    ByVal DecimalMode As Long _
) As Long
    On Error Resume Next
    Dim Opts As VnEngineOptions
    Opts.ZeroStyle = ZeroStyle
    Opts.ThousandStyle = ThousandStyle
    Opts.FourStyle = FourStyle
    Opts.DecimalMode = DecimalMode
    Dim res As String
    res = NumberToVietnamese(Value, Opts)
    GetNumberToVietnameseErrorCode = Err.Number
    On Error GoTo 0
End Function

Public Function IsValidNumberBridge(ByVal Value As Variant, ByRef OutErrorMessage As String) As Boolean
    IsValidNumberBridge = IsValidNumber(Value, OutErrorMessage)
End Function

' ------------------------------------------------------------------------------
' In-VBA Automated Core Test Runner
' ------------------------------------------------------------------------------

Public Function RunCoreTests() As String
    m_PassedCount = 0
    m_FailedCount = 0
    m_FirstFailure = ""

    ' 1. Single Digits
    AssertEqual "AC-01-01", NumberToVietnameseDefault(0), WordKhong()
    AssertEqual "AC-02-01", NumberToVietnameseDefault(1), WordMot()
    AssertEqual "AC-02-02", NumberToVietnameseDefault(5), WordNam()

    ' 2. Tens
    AssertEqual "AC-02-03", NumberToVietnameseDefault(10), WordMuoi()
    AssertEqual "AC-02-04", NumberToVietnameseDefault(11), WordMuoi() & " " & WordMot()
    AssertEqual "AC-02-05", NumberToVietnameseDefault(15), WordMuoi() & " " & WordLam()
    AssertEqual "AC-02-06", NumberToVietnameseDefault(20), WordHai() & " " & WordMuoiChuc()
    AssertEqual "AC-02-07", NumberToVietnameseDefault(21), WordHai() & " " & WordMuoiChuc() & " " & WordMotCuoi()
    AssertEqual "AC-02-08", NumberToVietnameseDefault(24), WordHai() & " " & WordMuoiChuc() & " " & WordTu()
    AssertEqual "AC-02-09", NumberToVietnameseDefault(25), WordHai() & " " & WordMuoiChuc() & " " & WordLam()
    AssertEqual "AC-02-10", NumberToVietnameseDefault(30), WordBa() & " " & WordMuoiChuc()
    AssertEqual "AC-02-11", NumberToVietnameseDefault(99), WordChin() & " " & WordMuoiChuc() & " " & WordChin()

    ' 3. Hundreds
    AssertEqual "AC-03-01", NumberToVietnameseDefault(100), WordMot() & " " & WordTram()
    AssertEqual "AC-03-02", NumberToVietnameseDefault(101), WordMot() & " " & WordTram() & " " & WordLe() & " " & WordMot()
    AssertEqual "AC-03-03", NumberToVietnameseDefault(105), WordMot() & " " & WordTram() & " " & WordLe() & " " & WordNam()
    AssertEqual "AC-03-04", NumberToVietnameseDefault(110), WordMot() & " " & WordTram() & " " & WordMuoi()
    AssertEqual "AC-03-05", NumberToVietnameseDefault(115), WordMot() & " " & WordTram() & " " & WordMuoi() & " " & WordLam()
    AssertEqual "AC-03-06", NumberToVietnameseDefault(121), WordMot() & " " & WordTram() & " " & WordHai() & " " & WordMuoiChuc() & " " & WordMotCuoi()
    AssertEqual "AC-03-07", NumberToVietnameseDefault(124), WordMot() & " " & WordTram() & " " & WordHai() & " " & WordMuoiChuc() & " " & WordTu()
    AssertEqual "AC-03-08", NumberToVietnameseDefault(125), WordMot() & " " & WordTram() & " " & WordHai() & " " & WordMuoiChuc() & " " & WordLam()
    AssertEqual "AC-03-09", NumberToVietnameseDefault(999), WordChin() & " " & WordTram() & " " & WordChin() & " " & WordMuoiChuc() & " " & WordChin()

    ' 4. Scale 10^3 (nghin)
    AssertEqual "AC-04-01", NumberToVietnameseDefault(1000), WordMot() & " " & WordNghin()
    AssertEqual "AC-04-02", NumberToVietnameseDefault(1001), WordMot() & " " & WordNghin() & " " & WordKhong() & " " & WordTram() & " " & WordLe() & " " & WordMot()
    AssertEqual "AC-04-03", NumberToVietnameseDefault(1005), WordMot() & " " & WordNghin() & " " & WordKhong() & " " & WordTram() & " " & WordLe() & " " & WordNam()
    AssertEqual "AC-04-04", NumberToVietnameseDefault(1010), WordMot() & " " & WordNghin() & " " & WordKhong() & " " & WordTram() & " " & WordMuoi()
    AssertEqual "AC-04-05", NumberToVietnameseDefault(1100), WordMot() & " " & WordNghin() & " " & WordMot() & " " & WordTram()
    AssertEqual "AC-04-06", NumberToVietnameseDefault(10000), WordMuoi() & " " & WordNghin()
    AssertEqual "AC-04-07", NumberToVietnameseDefault(100000), WordMot() & " " & WordTram() & " " & WordNghin()

    ' 5. Scale 10^6 (trieu)
    AssertEqual "AC-05-01", NumberToVietnameseDefault(1000000), WordMot() & " " & WordTrieu()
    AssertEqual "AC-05-02", NumberToVietnameseDefault(1000001), WordMot() & " " & WordTrieu() & " " & WordKhong() & " " & WordTram() & " " & WordLe() & " " & WordMot()
    AssertEqual "AC-05-03", NumberToVietnameseDefault(1000005), WordMot() & " " & WordTrieu() & " " & WordKhong() & " " & WordTram() & " " & WordLe() & " " & WordNam()
    AssertEqual "AC-05-04", NumberToVietnameseDefault(1000010), WordMot() & " " & WordTrieu() & " " & WordKhong() & " " & WordTram() & " " & WordMuoi()
    AssertEqual "AC-05-05", NumberToVietnameseDefault(1000100), WordMot() & " " & WordTrieu() & " " & WordMot() & " " & WordTram()
    AssertEqual "AC-05-06", NumberToVietnameseDefault(1001000), WordMot() & " " & WordTrieu() & " " & WordKhong() & " " & WordTram() & " " & WordLe() & " " & WordMot() & " " & WordNghin()
    AssertEqual "AC-05-08", NumberToVietnameseDefault(1050000), WordMot() & " " & WordTrieu() & " " & WordKhong() & " " & WordTram() & " " & WordNam() & " " & WordMuoiChuc() & " " & WordNghin()

    ' 6. Scale 10^9 (ty) and 10^12 (nghin ty)
    AssertEqual "AC-06-01", NumberToVietnameseDefault(1000000000#), WordMot() & " " & WordTy()
    AssertEqual "AC-06-02", NumberToVietnameseDefault(1000000001#), WordMot() & " " & WordTy() & " " & WordKhong() & " " & WordTram() & " " & WordLe() & " " & WordMot()
    AssertEqual "AC-07-01", NumberToVietnameseDefault(1000000000000#), WordMot() & " " & WordNghin() & " " & WordTy()
    AssertEqual "AC-07-02", NumberToVietnameseDefault(1000000000001#), WordMot() & " " & WordNghin() & " " & WordTy() & " " & WordKhong() & " " & WordTram() & " " & WordLe() & " " & WordMot()

    ' 7. Boundaries
    Dim end8 As String
    end8 = WordChin() & " " & WordTram() & " " & WordChin() & " " & WordMuoiChuc() & " " & WordTam()
    AssertEqual "AC-08-01", Right$(NumberToVietnameseDefault(999999999999998#), Len(end8)), end8

    Dim end9 As String
    end9 = WordChin() & " " & WordTram() & " " & WordChin() & " " & WordMuoiChuc() & " " & WordChin()
    AssertEqual "AC-08-02", Right$(NumberToVietnameseDefault(999999999999999#), Len(end9)), end9

    ' 8. Negatives
    AssertEqual "AC-10-01", NumberToVietnameseDefault(-100), ToUnicodeLower(WordAm()) & " " & WordMot() & " " & WordTram()
    AssertEqual "AC-10-02", NumberToVietnameseDefault(-150000), ToUnicodeLower(WordAm()) & " " & WordMot() & " " & WordTram() & " " & WordNam() & " " & WordMuoiChuc() & " " & WordNghin()
    AssertEqual "AC-10-03", NumberToVietnameseDefault(-1000000), ToUnicodeLower(WordAm()) & " " & WordMot() & " " & WordTrieu()

    ' 9. Decimals
    AssertEqual "AC-11-01", NumberToVietnameseWithOptions(125.5, 0, 0, 0, 0), WordMot() & " " & WordTram() & " " & WordHai() & " " & WordMuoiChuc() & " " & WordLam()
    AssertEqual "AC-11-02", NumberToVietnameseWithOptions(125.05, 0, 0, 0, 1), WordMot() & " " & WordTram() & " " & WordHai() & " " & WordMuoiChuc() & " " & WordLam() & " " & WordPhay() & " " & WordKhong() & " " & WordNam()
    AssertEqual "AC-11-03", NumberToVietnameseWithOptions(-125.9, 0, 0, 0, 0), ToUnicodeLower(WordAm()) & " " & WordMot() & " " & WordTram() & " " & WordHai() & " " & WordMuoiChuc() & " " & WordLam()
    AssertEqual "AC-11-04", NumberToVietnameseWithOptions(-125.9, 0, 0, 0, 1), ToUnicodeLower(WordAm()) & " " & WordMot() & " " & WordTram() & " " & WordHai() & " " & WordMuoiChuc() & " " & WordLam() & " " & WordPhay() & " " & WordChin()

    ' 10. Dialects
    AssertEqual "AC-12-01", NumberToVietnameseWithOptions(105, 1, 0, 0, 0), WordMot() & " " & WordTram() & " " & WordLinh() & " " & WordNam()
    AssertEqual "AC-12-03", NumberToVietnameseWithOptions(1000, 0, 1, 0, 0), WordMot() & " " & WordNgan()
    AssertEqual "AC-12-05", NumberToVietnameseWithOptions(24, 0, 0, 1, 0), WordHai() & " " & WordMuoiChuc() & " " & WordBon()
    AssertEqual "AC-12-07", NumberToVietnameseWithOptions(1005, 1, 1, 0, 0), WordMot() & " " & WordNgan() & " " & WordKhong() & " " & WordTram() & " " & WordLinh() & " " & WordNam()

    ' 11. Error assertions
    AssertRaisesError "AC-09-01", 1000000000000000#, ERR_OUT_OF_RANGE
    AssertRaisesError "AC-09-03", "abc", ERR_INVALID_NUMBER

    If m_FailedCount = 0 Then
        RunCoreTests = "PASS: " & CStr(m_PassedCount) & "/" & CStr(m_PassedCount + m_FailedCount) & " assertions passed."
    Else
        RunCoreTests = "FAIL: " & CStr(m_FailedCount) & " failures. First: " & m_FirstFailure
    End If
End Function

Private Sub AssertEqual(ByVal CaseId As String, ByVal Actual As String, ByVal Expected As String)
    If Actual = Expected Then
        m_PassedCount = m_PassedCount + 1
    Else
        m_FailedCount = m_FailedCount + 1
        If Len(m_FirstFailure) = 0 Then
            m_FirstFailure = "[" & CaseId & "] Expected '" & Expected & "', got '" & Actual & "'"
        End If
    End If
End Sub

Private Sub AssertRaisesError(ByVal CaseId As String, ByVal Value As Variant, ByVal ExpectedErr As Long)
    On Error Resume Next
    Dim Opts As VnEngineOptions
    Opts = DefaultEngineOptions()
    Dim res As String
    res = NumberToVietnamese(Value, Opts)
    Dim errNum As Long
    errNum = Err.Number
    On Error GoTo 0
    
    If errNum = ExpectedErr Then
        m_PassedCount = m_PassedCount + 1
    Else
        m_FailedCount = m_FailedCount + 1
        If Len(m_FirstFailure) = 0 Then
            m_FirstFailure = "[" & CaseId & "] Expected error " & CStr(ExpectedErr) & ", got " & CStr(errNum)
        End If
    End If
End Sub
