Attribute VB_Name = "TextFormatter"
Option Explicit

' ==============================================================================
' TextFormatter.bas
' Pure VBA text normalization, whitespace collapsing, Unicode casing, punctuation.
' Zero Excel dependencies, zero UI, zero external DLL dependencies.
' All non-ASCII characters accessed programmatically via UnicodeText.bas.
' Source code is 100% 7-bit ASCII safe.
' ==============================================================================

Public Function FormatText( _
    ByVal Text As String, _
    Optional ByVal Casing As VnCasingStyle = VnCaseSentence, _
    Optional ByVal AddPeriod As Boolean = True _
) As String
    If Len(Text) = 0 Then
        FormatText = ""
        Exit Function
    End If
    
    ' 1. Whitespace normalization: convert tabs/CR/LF to space, collapse runs, trim
    Dim s As String
    s = Text
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Trim$(s)
    
    If Len(s) = 0 Then
        FormatText = ""
        Exit Function
    End If
    
    ' 2. Remove whitespace before punctuation
    s = Replace(s, " .", ".")
    s = Replace(s, " ,", ",")
    s = Replace(s, " ;", ";")
    s = Replace(s, " :", ":")
    s = Replace(s, " !", "!")
    s = Replace(s, " ?", "?")
    
    ' 3. Apply casing via UnicodeText
    Select Case Casing
        Case VnCaseSentence
            s = CapitalizeFirst(s)
        Case VnCaseUpper
            s = ToUnicodeUpper(s)
        Case VnCaseLower
            s = ToUnicodeLower(s)
        Case Else
            Err.Raise ERR_INVALID_OPTIONS, "TextFormatter.FormatText", "Tuy chon casing khong hop le."
    End Select
    
    ' 4. Append trailing period if AddPeriod = True (ensure no double periods)
    If AddPeriod Then
        If Right$(s, 1) <> "." Then
            s = s & "."
        End If
    End If
    
    FormatText = s
End Function
