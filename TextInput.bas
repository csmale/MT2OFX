Attribute VB_Name = "TextInput"
Option Explicit

' $Header: /MT2OFX/TextInput.bas 15    6/12/14 22:03 Colin $

#Const debugmode = 0
#Const OldVersion = False

'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       UnquotedItem
' Description:       Returns next part of string up to delimiter
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       21/11/2003-21:43:56
'
' Parameters :       sLine (String)
'                    Delimiter (String)
'--------------------------------------------------------------------------------
'</CSCM>
Private Function UnquotedItem(sLine As String, Delimiter As String, bUseBackslash As Boolean) As String
    Dim sTmp As String
    Dim c As String
    Dim i As Integer
    If Len(sLine) = 0 Then
        UnquotedItem = ""
        Exit Function
    End If
    i = 1
    c = Left$(sLine, 1)
    Do While c <> Delimiter
        If bUseBackslash Then
            If c = "\" Then
                If i >= Len(sLine) Then Exit Do
                i = i + 1
                c = Mid$(sLine, i, 1)
            End If
        End If
        sTmp = sTmp + c
        If i >= Len(sLine) Then Exit Do
        i = i + 1
        c = Mid$(sLine, i, 1)
    Loop
    sLine = Mid$(sLine, i + 1)
    UnquotedItem = sTmp
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       QuotedItem
' Description:       Returns next part of string up to terminating quote
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       21/11/2003-21:44:58
'
' Parameters :       sLine (String)
'--------------------------------------------------------------------------------
'</CSCM>
Private Function QuotedItem(sLine As String, Delimiter As String, bUseBackslash As Boolean) As String
    Dim sTmp As String
    Dim c As String
    Dim q As String
    Dim i As Integer
    If Len(sLine) = 0 Then
        QuotedItem = ""
        Exit Function
    End If
    i = 2
    q = Left$(sLine, 1)
    If q <> """" And q <> "'" Then
        QuotedItem = UnquotedItem(sLine, Delimiter, bUseBackslash)
        Exit Function
    End If
tryagain:
    c = Mid$(sLine, i, 1)
    Do While c <> q
        If bUseBackslash Then
            If c = "\" Then
                If i >= Len(sLine) Then Exit Do
                i = i + 1
                c = Mid$(sLine, i, 1)
            End If
        End If
        sTmp = sTmp + c
        i = i + 1
        If i > Len(sLine) Then Exit Do
        c = Mid$(sLine, i, 1)
    Loop
    If i >= Len(sLine) Then GoTo goback
    If Mid$(sLine, i + 1, 1) <> Delimiter Then
        If Mid$(sLine, i + 1, 1) = q Then  ' embedded doubled quote - acts as single
            i = i + 1 ' skip extra char
        End If
        sTmp = sTmp & c ' tack on the spurious quote
        i = i + 1
        GoTo tryagain
    End If
goback:
    sLine = Mid$(sLine, i + 2)
    QuotedItem = sTmp
End Function

'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtParseLineDelimited
' Description:       Parse a line into fields using delimiter
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       05/11/2003-21:51:51
'
' Parameters :       Line (String)
'                    Delimiter (String), typically ","
'                    Handles quoted strings (sq or dq)
'                    Handles escaped special chars (with backslash)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function txtParseLineDelimited(ByVal Line As String, Delimiter As String, bUseBackslash As Boolean) As Variant
    Dim vFields As Variant
    Dim iCount As Integer
    Dim bEndsWithDelim As Boolean
    bEndsWithDelim = (Right$(Line, 1) = Delimiter)
    iCount = 0
    ReDim vFields(1 To 1000)
    Dim sTmp As String
    Dim iPos As Long
    Do While Len(Line) > 0
        sTmp = QuotedItem(Line, Delimiter, bUseBackslash)
        iCount = iCount + 1
'        ReDim Preserve vFields(1 To iCount)
        vFields(iCount) = sTmp
    Loop
    If bEndsWithDelim Then
        iCount = iCount + 1
'        ReDim Preserve vFields(1 To iCount)
        vFields(iCount) = ""
    End If
    If iCount = 0 Then
        vFields = Array()
    Else
        ReDim Preserve vFields(1 To iCount)
    End If
    txtParseLineDelimited = vFields
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       MakeLineDelimited
' Description:       Create CSV-compatible line from items
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       30/07/2004-23:32:57
'
' Parameters :       vFields (Variant)
'                    Delimiter (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function MakeLineDelimited(vFields As Variant, Delimiter As String) As String
    Dim sTmp As String
    Dim sLine As String
    Dim i As Long
    If IsArray(vFields) Then
        For i = LBound(vFields) To UBound(vFields)
            sTmp = CStr(vFields(i))
            sTmp = Replace(sTmp, """", """""")  ' double quotes - doubled up!
            If Len(sLine) > 0 Then
                sLine = sLine & Delimiter
            End If
            sLine = sLine & """" & sTmp & """"
        Next
    Else
        sTmp = CStr(vFields)
        sTmp = Replace(sTmp, """", """""")  ' double quotes - doubled up!
        sLine = """" & sTmp & """"
    End If
    MakeLineDelimited = sLine
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtParseLineFixed
' Description:       Parse line as fixed width fields
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       05/11/2003-21:53:39
'
' Parameters :       Line (String)
'                    Pattern (String)
'                    Pattern is for RegExp with fields between parens
'                    Eg. "(\d{10})([A-Z]{3})" matches 10 digits followed by 3 letters
'--------------------------------------------------------------------------------
'</CSCM>
Public Function txtParseLineFixed(ByVal Line As String, ByVal Pattern As String) As Variant
    Dim vFields As Variant
    Dim iCount As Integer
#If debugmode > 0 Then
    LogMessage False, True, "Enter: txtParseLineFixed"
#End If
    iCount = 0
    Dim r As New RegExp
    r.Pattern = "^" & Pattern & "$"
#If debugmode > 0 Then
    LogMessage False, True, "txtParseLineFixed: RegExp created"
#End If
    Dim res As MatchCollection
    Set res = r.Execute(Line)
#If debugmode > 0 Then
    LogMessage False, True, "txtParseLineFixed: Match executed"
#End If
    If res.Count = 0 Then
#If debugmode > 0 Then
    LogMessage False, True, "txtParseLineFixed: No Match!"
#End If
        txtParseLineFixed = Array()
        Exit Function
    End If
#If debugmode > 0 Then
    LogMessage False, True, "txtParseLineFixed: match complete, returning fields"
#End If
    Dim i As Long
    i = 0
    ReDim vFields(1 To res(0).SubMatches.Count)
    For i = 1 To res(0).SubMatches.Count
        vFields(i) = res(0).SubMatches(i - 1)
    Next
    txtParseLineFixed = vFields
#If debugmode > 0 Then
    LogMessage False, True, "Leave: txtParseLineFixed"
#End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       NormaliseLineEndings
' Description:       Change all kinds of line endings into CR
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       2/6/2008-21:54:11
'
' Parameters :       sIn (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function NormaliseLineEndings(sIn As String) As String
' sort out line terminators
    Dim sTmp As String
#If False Then
    Dim iCR: iCR = InStr(sIn, vbCr)
    Dim iLF: iLF = InStr(sIn, vbLf)
    
    If iLF = 0 Then
        If iCR > 0 Then
            sTmp = sIn  ' we are done!
        Else
            ' no lf or cr...?
' Dortmunder Stadtsparkasse uses "@@" to separate the lines!!!???
            sTmp = Replace$(sIn, "@@", vbCr)
        End If
    Else
        If iCR = 0 Then ' lf, no cr
            sTmp = Replace$(sIn, vbLf, vbCr)
        Else    ' lf and cr - remove the lf's
' problem with Barclays International: start of file is CR+LF, thereafter only LF is used. So we make sure there are CR's
' in the rest of the file! If they are there, use them and drop the LFs. Otherwise (no further CRs) we convert the LFs to CRs
            If InStr(iLF + 1, sIn, vbCr) > 0 Then
                sTmp = Replace$(sIn, vbLf, "")
            Else
                sTmp = Replace$(sIn, vbLf, vbCr)
            End If
        End If
    End If
    NormaliseLineEndings = sTmp
    Exit Function
#Else
    Dim sLineEnd As String
    sLineEnd = ChrW(&HE123) ' range 0xe000-0xf8ff are unused in unicode
    sTmp = Replace$(sIn, vbCr & vbLf, sLineEnd)
    sTmp = Replace$(sTmp, vbLf & vbCr, sLineEnd)
    sTmp = Replace$(sTmp, vbCr, sLineEnd)
    sTmp = Replace$(sTmp, vbLf, sLineEnd)
' Dortmunder Stadtsparkasse uses "@@" to separate the lines!!!???
    sTmp = Replace$(sTmp, "@@", sLineEnd)
    NormaliseLineEndings = Replace$(sTmp, sLineEnd, vbCr)
#End If
End Function
