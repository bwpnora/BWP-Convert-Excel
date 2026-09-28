Attribute VB_Name = "UnicodeText"
Option Explicit

' ==============================================================================
' UnicodeText.bas
' ASCII-safe Vietnamese vocabulary and complete Unicode character casing engine.
' Pure VBA: Zero Excel dependencies, zero external DLL dependencies.
' All non-ASCII characters generated programmatically via ChrW$(&H...).
' ==============================================================================

Private m_Initialized As Boolean
Private m_UpperMap(0 To 8191) As Integer
Private m_LowerMap(0 To 8191) As Integer

' ------------------------------------------------------------------------------
' Vietnamese Core Vocabulary
' ------------------------------------------------------------------------------

Public Function WordKhong() As String
    WordKhong = "kh" & ChrW$(&H00F4) & "ng"
End Function

Public Function WordMot() As String
    WordMot = "m" & ChrW$(&H1ED9) & "t"
End Function

Public Function WordHai() As String
    WordHai = "hai"
End Function

Public Function WordBa() As String
    WordBa = "ba"
End Function

Public Function WordBon() As String
    WordBon = "b" & ChrW$(&H1ED1) & "n"
End Function

Public Function WordNam() As String
    WordNam = "n" & ChrW$(&H0103) & "m"
End Function

Public Function WordSau() As String
    WordSau = "s" & ChrW$(&H00E1) & "u"
End Function

Public Function WordBay() As String
    WordBay = "b" & ChrW$(&H1EA3) & "y"
End Function

Public Function WordTam() As String
    WordTam = "t" & ChrW$(&H00E1) & "m"
End Function

Public Function WordChin() As String
    WordChin = "ch" & ChrW$(&H00ED) & "n"
End Function

Public Function WordMuoi() As String
    WordMuoi = "m" & ChrW$(&H01B0) & ChrW$(&H1EDD) & "i"
End Function

Public Function WordMuoiChuc() As String
    WordMuoiChuc = "m" & ChrW$(&H01B0) & ChrW$(&H01A1) & "i"
End Function

Public Function WordTram() As String
    WordTram = "tr" & ChrW$(&H0103) & "m"
End Function

Public Function WordNghin() As String
    WordNghin = "ngh" & ChrW$(&H00EC) & "n"
End Function

Public Function WordNgan() As String
    WordNgan = "ng" & ChrW$(&H00E0) & "n"
End Function

Public Function WordTrieu() As String
    WordTrieu = "tri" & ChrW$(&H1EC7) & "u"
End Function

Public Function WordTy() As String
    WordTy = "t" & ChrW$(&H1EF7)
End Function

Public Function WordLe() As String
    WordLe = "l" & ChrW$(&H1EBB)
End Function

Public Function WordLinh() As String
    WordLinh = "linh"
End Function

Public Function WordMotCuoi() As String
    WordMotCuoi = "m" & ChrW$(&H1ED1) & "t"
End Function

Public Function WordTu() As String
    WordTu = "t" & ChrW$(&H01B0)
End Function

Public Function WordLam() As String
    WordLam = "l" & ChrW$(&H0103) & "m"
End Function

Public Function WordAm() As String
    WordAm = ChrW$(&H00C2) & "m"
End Function

Public Function WordPhay() As String
    WordPhay = "ph" & ChrW$(&H1EA9) & "y"
End Function

Public Function WordDong() As String
    WordDong = ChrW$(&H0111) & ChrW$(&H1ED3) & "ng"
End Function

Public Function WordDongHoa() As String
    WordDongHoa = ChrW$(&H0110) & ChrW$(&H1ED3) & "ng"
End Function

Public Function WordChan() As String
    WordChan = "ch" & ChrW$(&H1EB5) & "n"
End Function

Public Function WordDolaMy() As String
    WordDolaMy = ChrW$(&H0111) & ChrW$(&H00F4) & " la M" & ChrW$(&H1EF9)
End Function

Public Function WordCent() As String
    WordCent = "cent"
End Function

Public Function WordXu() As String
    WordXu = "xu"
End Function

' ------------------------------------------------------------------------------
' Unicode Casing Tables & Functions
' ------------------------------------------------------------------------------

Private Sub MapPair(ByVal LowerCode As Long, ByVal UpperCode As Long)
    m_UpperMap(LowerCode) = CInt(UpperCode)
    m_LowerMap(UpperCode) = CInt(LowerCode)
End Sub

Private Sub EnsureCaseMaps()
    If m_Initialized Then Exit Sub

    Dim i As Long
    For i = 0 To 8191
        m_UpperMap(i) = CInt(i)
        m_LowerMap(i) = CInt(i)
    Next i

    ' Standard ASCII Latin letters (A-Z / a-z)
    For i = 97 To 122 ' a-z -> A-Z
        MapPair i, i - 32
    Next i

    ' Vietnamese Accented Vowels and Consonants (67 pairs / 134 characters)
    ' A with diacritics
    MapPair &H00E0, &H00C0 ' a with grave / A with grave
    MapPair &H00E1, &H00C1 ' a with acute / A with acute
    MapPair &H1EA3, &H1EA2 ' a with hook above / A with hook above
    MapPair &H00E3, &H00C3 ' a with tilde / A with tilde
    MapPair &H1EA1, &H1EA0 ' a with dot below / A with dot below
    MapPair &H0103, &H0102 ' a with breve / A with breve
    MapPair &H1EB1, &H1EB0 ' a with breve and grave / A with breve and grave
    MapPair &H1EAF, &H1EAE ' a with breve and acute / A with breve and acute
    MapPair &H1EB3, &H1EB2 ' a with breve and hook above / A with breve and hook above
    MapPair &H1EB5, &H1EB4 ' a with breve and tilde / A with breve and tilde
    MapPair &H1EB7, &H1EB6 ' a with breve and dot below / A with breve and dot below
    MapPair &H00E2, &H00C2 ' a with circumflex / A with circumflex
    MapPair &H1EA7, &H1EA6 ' a with circumflex and grave / A with circumflex and grave
    MapPair &H1EA5, &H1EA4 ' a with circumflex and acute / A with circumflex and acute
    MapPair &H1EA9, &H1EA8 ' a with circumflex and hook above / A with circumflex and hook above
    MapPair &H1EAB, &H1EAA ' a with circumflex and tilde / A with circumflex and tilde
    MapPair &H1EAD, &H1EAC ' a with circumflex and dot below / A with circumflex and dot below

    ' E with diacritics
    MapPair &H00E8, &H00C8 ' e with grave / E with grave
    MapPair &H00E9, &H00C9 ' e with acute / E with acute
    MapPair &H1EBB, &H1EBA ' e with hook above / E with hook above
    MapPair &H1EBD, &H1EBC ' e with tilde / E with tilde
    MapPair &H1EB9, &H1EB8 ' e with dot below / E with dot below
    MapPair &H00EA, &H00CA ' e with circumflex / E with circumflex
    MapPair &H1EC1, &H1EC0 ' e with circumflex and grave / E with circumflex and grave
    MapPair &H1EBF, &H1EBE ' e with circumflex and acute / E with circumflex and acute
    MapPair &H1EC3, &H1EC2 ' e with circumflex and hook above / E with circumflex and hook above
    MapPair &H1EC5, &H1EC4 ' e with circumflex and tilde / E with circumflex and tilde
    MapPair &H1EC7, &H1EC6 ' e with circumflex and dot below / E with circumflex and dot below

    ' I with diacritics
    MapPair &H00EC, &H00CC ' i with grave / I with grave
    MapPair &H00ED, &H00CD ' i with acute / I with acute
    MapPair &H1EC9, &H1EC8 ' i with hook above / I with hook above
    MapPair &H0129, &H0128 ' i with tilde / I with tilde
    MapPair &H1ECB, &H1ECA ' i with dot below / I with dot below

    ' O with diacritics
    MapPair &H00F2, &H00D2 ' o with grave / O with grave
    MapPair &H00F3, &H00D3 ' o with acute / O with acute
    MapPair &H1ECF, &H1ECE ' o with hook above / O with hook above
    MapPair &H00F5, &H00D5 ' o with tilde / O with tilde
    MapPair &H1ECD, &H1ECC ' o with dot below / O with dot below
    MapPair &H00F4, &H00D4 ' o with circumflex / O with circumflex
    MapPair &H1ED3, &H1ED2 ' o with circumflex and grave / O with circumflex and grave
    MapPair &H1ED1, &H1ED0 ' o with circumflex and acute / O with circumflex and acute
    MapPair &H1ED5, &H1ED4 ' o with circumflex and hook above / O with circumflex and hook above
    MapPair &H1ED7, &H1ED6 ' o with circumflex and tilde / O with circumflex and tilde
    MapPair &H1ED9, &H1ED8 ' o with circumflex and dot below / O with circumflex and dot below
    MapPair &H01A1, &H01A0 ' o with horn / O with horn
    MapPair &H1EDD, &H1EDC ' o with horn and grave / O with horn and grave
    MapPair &H1EDB, &H1EDA ' o with horn and acute / O with horn and acute
    MapPair &H1EDF, &H1EDE ' o with horn and hook above / O with horn and hook above
    MapPair &H1EE1, &H1EE0 ' o with horn and tilde / O with horn and tilde
    MapPair &H1EE3, &H1EE2 ' o with horn and dot below / O with horn and dot below

    ' U with diacritics
    MapPair &H00F9, &H00D9 ' u with grave / U with grave
    MapPair &H00FA, &H00DA ' u with acute / U with acute
    MapPair &H1EE7, &H1EE6 ' u with hook above / U with hook above
    MapPair &H0169, &H0168 ' u with tilde / U with tilde
    MapPair &H1EE5, &H1EE4 ' u with dot below / U with dot below
    MapPair &H01B0, &H01AF ' u with horn / U with horn
    MapPair &H1EEB, &H1EEA ' u with horn and grave / U with horn and grave
    MapPair &H1EE9, &H1EE8 ' u with horn and acute / U with horn and acute
    MapPair &H1EED, &H1EEC ' u with horn and hook above / U with horn and hook above
    MapPair &H1EEF, &H1EEE ' u with horn and tilde / U with horn and tilde
    MapPair &H1EF1, &H1EF0 ' u with horn and dot below / U with horn and dot below

    ' Y with diacritics
    MapPair &H1EF3, &H1EF2 ' y with grave / Y with grave
    MapPair &H00FD, &H00DD ' y with acute / Y with acute
    MapPair &H1EF7, &H1EF6 ' y with hook above / Y with hook above
    MapPair &H1EF9, &H1EF8 ' y with tilde / Y with tilde
    MapPair &H1EF5, &H1EF4 ' y with dot below / Y with dot below

    ' D with stroke
    MapPair &H0111, &H0110 ' d with stroke / D with stroke

    m_Initialized = True
End Sub

Public Function ToUnicodeUpper(ByVal Text As String) As String
    If Len(Text) = 0 Then
        ToUnicodeUpper = ""
        Exit Function
    End If

    EnsureCaseMaps
    Dim result As String
    result = Text

    Dim i As Long
    Dim code As Long
    Dim mapCode As Integer

    For i = 1 To Len(result)
        code = AscW(Mid$(result, i, 1))
        If code >= 0 And code <= 8191 Then
            mapCode = m_UpperMap(code)
            If mapCode <> code Then
                Mid$(result, i, 1) = ChrW$(mapCode)
            End If
        End If
    Next i

    ToUnicodeUpper = result
End Function

Public Function ToUnicodeLower(ByVal Text As String) As String
    If Len(Text) = 0 Then
        ToUnicodeLower = ""
        Exit Function
    End If

    EnsureCaseMaps
    Dim result As String
    result = Text

    Dim i As Long
    Dim code As Long
    Dim mapCode As Integer

    For i = 1 To Len(result)
        code = AscW(Mid$(result, i, 1))
        If code >= 0 And code <= 8191 Then
            mapCode = m_LowerMap(code)
            If mapCode <> code Then
                Mid$(result, i, 1) = ChrW$(mapCode)
            End If
        End If
    Next i

    ToUnicodeLower = result
End Function

Public Function CapitalizeFirst(ByVal Text As String) As String
    If Len(Text) = 0 Then
        CapitalizeFirst = ""
        Exit Function
    End If

    EnsureCaseMaps
    Dim result As String
    result = Text

    Dim i As Long
    Dim ch As String
    Dim code As Long

    For i = 1 To Len(result)
        ch = Mid$(result, i, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then
            code = AscW(ch)
            If code >= 0 And code <= 8191 Then
                If m_UpperMap(code) <> code Then
                    Mid$(result, i, 1) = ChrW$(m_UpperMap(code))
                End If
            End If
            Exit For
        End If
    Next i

    CapitalizeFirst = result
End Function
