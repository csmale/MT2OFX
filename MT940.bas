Attribute VB_Name = "MT940"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : MT940
'    Project    : MT2OFX
'
'    Description:
'
'    Modified   : $Author: Colin $ $Date: 15/06/09 19:25 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/MT940.bas 17    15/06/09 19:25 Colin $"
' $History: MT940.bas $
' 
' *****************  Version 17  *****************
' User: Colin        Date: 15/06/09   Time: 19:25
' Updated in $/MT2OFX
' For transfer to new laptop
'
' *****************  Version 15  *****************
' User: Colin        Date: 19/04/08   Time: 22:14
' Updated in $/MT2OFX
' changed to work with new i/o system
'
' *****************  Version 14  *****************
' User: Colin        Date: 7/12/06    Time: 15:07
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 11  *****************
' User: Colin        Date: 2/11/05    Time: 23:03
' Updated in $/MT2OFX
' V3.4 beta 1
'
' *****************  Version 10  *****************
' User: Colin        Date: 11/06/05   Time: 19:33
' Updated in $/MT2OFX
'
' *****************  Version 9  *****************
' User: Colin        Date: 18/03/05   Time: 21:57
' Updated in $/MT2OFX
'</CSCC>

Public Session As New Session
Private iLineNum As Long
Private sLastLine As String
Private sPrevLine As String
Private sLastColonLine As String
Private sPrevColonLine As String

Dim sTxn As String
Dim sText As String

Private Function ParseDate6(DateString As String) As Date
' 6-digit date, YYMMDD
    If Not IsNumeric(DateString) Or Len(DateString) < 6 Then
        LogMessage True, True, "ParseDate6: Unable to parse '" & DateString & "' as a date."
        LogMessage False, True, "Near line #" & CStr(iLineNum - 1) & ": '" & sPrevColonLine & "'"
        ParseDate6 = NODATE
        Exit Function
    End If
    Dim iYear As Integer
    iYear = CInt(Mid$(DateString, 1, 2)) + 1900
    If iYear < 1980 Then iYear = iYear + 100
    ParseDate6 = DateSerial(iYear, _
        CInt(Mid$(DateString, 3, 2)), _
        CInt(Mid$(DateString, 5, 2)))
End Function
Private Function ParseDate4(DateString As String, BaseDate As Date) As Date
' 4-digit date, MMDD
    If Not IsNumeric(DateString) Or Len(DateString) < 4 Then
        LogMessage True, True, "ParseDate4: Unable to parse '" & DateString & "' as a date."
        LogMessage False, True, "Near line #" & CStr(iLineNum - 1) & ": '" & sPrevColonLine & "'"
        ParseDate4 = NODATE
        Exit Function
    End If
    Dim iYear As Integer
    iYear = Year(BaseDate)
    ' do we need to watch out for year-end rollover or not?
    ' do it anyway. it'll only be a few days anyway.
    If Month(BaseDate) = 12 And CInt(Mid$(DateString, 1, 2)) = 1 Then
        iYear = iYear + 1
    ElseIf Month(BaseDate) = 1 And CInt(Mid$(DateString, 1, 2)) = 12 Then
        iYear = iYear - 1
    End If
    ParseDate4 = DateSerial(iYear, _
        CInt(Mid$(DateString, 1, 2)), _
        CInt(Mid$(DateString, 3, 2)))
End Function
Private Function ParseBalance(b As Balance, BalanceString As String) As Boolean
' e.g. C021004EUR1339,51
    Dim iPos As Integer
    Dim iSign As Integer
    iSign = IIf(Left$(BalanceString, 1) = "C", 1, -1)
    Dim dBalDate As Date
    dBalDate = ParseDate6(Mid$(BalanceString, 2, 6))
    Dim sCurr As String
    If Not IsNumeric(Mid$(BalanceString, 8, 1)) Then
        sCurr = Mid$(BalanceString, 8, 3)
        iPos = 11
    Else
        sCurr = ""
        iPos = 8
    End If
    Dim sAmt As String
    sAmt = Mid$(BalanceString, iPos)
    Dim sPoint As String
' MT940 should always use a comma as decimal separator, whereal CDbl() uses the
' local settings
    If InStr(sAmt, ".") <> 0 Then
        sPoint = "."
    Else
        sPoint = ","
    End If
' MT940 always uses a comma as decimal separator, whereal CDbl() uses the
' local settings
    Dim dAmt As Double
    dAmt = NumberFromString(sAmt, sPoint) * iSign
    
    b.BalDate = dBalDate
    b.Amt = dAmt
    b.Ccy = sCurr
    
    ParseBalance = True
End Function

' this is for the :61: line with transaction details
Private Function ParseTransaction(t As Txn, TxnString As String) As Boolean
' e.g. 0211191119D21,31N426NONREF
    ParseTransaction = False
    On Error GoTo parseerr
    t.ValueDate = ParseDate6(Left$(TxnString, 6))
    If t.ValueDate = NODATE Then GoTo parseerr
    Dim iPos As Integer
    iPos = 7
    If IsNumeric(Mid$(TxnString, 7, 1)) Then
        t.BookDate = ParseDate4(Mid$(TxnString, 7, 4), t.ValueDate)
        If t.BookDate = NODATE Then GoTo parseerr
        iPos = 11
    Else
        t.BookDate = t.ValueDate
    End If
    Dim iSign As Integer
    If Mid$(TxnString, iPos) = "R" Then
        iPos = iPos + 1
        t.IsReversal = True
    Else
        t.IsReversal = False
    End If
    iSign = IIf(Mid$(TxnString, iPos, 1) = "C", 1, -1)
    iPos = iPos + 1
' it seems some banks use other codes than C/D/RC/RD ---
    If Not IsNumeric(Mid$(TxnString, iPos, 1)) Then
        iPos = iPos + 1
    End If
    Dim iTmp As Integer
    iTmp = iPos
    While InStr("0123456789,.", Mid$(TxnString, iTmp, 1)) <> 0
        iTmp = iTmp + 1
    Wend
    Dim sPoint As String
    Dim sAmt As String
    sAmt = Mid$(TxnString, iPos, iTmp - iPos)
' MT940 should always use a comma as decimal separator, whereal CDbl() uses the
' local settings
    If InStr(sAmt, ".") <> 0 Then
        sPoint = "."
    Else
        sPoint = ","
    End If
    t.Amt = NumberFromString(sAmt, sPoint) * iSign
    iPos = iTmp
    t.BookingCode = Mid$(TxnString, iPos, 4)
    iPos = iPos + 4
    t.Reference = Mid$(TxnString, iPos)
    iTmp = InStr(t.Reference, "//")
    If iTmp = 0 Then
        If Len(t.Reference) > 16 Then
            t.BankReference = Mid$(t.Reference, 17)
            t.Reference = Trim$(Left$(t.Reference, 16))
        Else
            t.BankReference = ""
        End If
    Else
        t.BankReference = Mid$(t.Reference, iTmp + 2)
        t.Reference = Left$(t.Reference, iTmp - 1)
    End If
    ParseTransaction = True
    Exit Function
parseerr:
    
End Function

Private Function FlushTxn(stmt As Statement, sTxn As String, _
    sText As String, t As Txn) As Boolean
    Dim sTmp As String
    If sTxn = "" Then
        FlushTxn = True
        Exit Function
    End If
    If Not ParseTransaction(t, sTxn) Then
        LogMessage True, True, "Unable to parse transaction: '" & sTxn & "'"
    Else
        If Bcfg.Structured86 Then
            If t.Str86.ParseInfo(sText) Then
                t.FurtherInfo = t.Str86.Memo
                sTmp = t.Str86.GetField(sfTextCodeSupplement)
                If sTmp <> "" Then t.BookingCode = sTmp
            Else
                t.FurtherInfo = sText
                t.BookingCode = -1
            End If
        Else
            t.FurtherInfo = sText
        End If
        t.Index = stmt.Txns.Count + 1
        stmt.Txns.Add t
        ' reset the txn
    End If
    sTxn = ""
    FlushTxn = True
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OtherLine
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       05/03/2004-22:46:58
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub OtherLine(sLine As String, sText As String)
    If Bcfg.Structured86 Then
        sText = sText & sLine
    Else
        If Bcfg.SkipEmptyMemoFields Then
            If Trim$(sLine) <> "" Then
                If sText = "" Then
                    sText = Trim$(sLine)
                Else
                    sText = sText & Cfg.MemoDelimiter & Trim$(sLine)
                End If
            End If
        Else
            sText = sText & Cfg.MemoDelimiter & Trim$(sLine)
        End If
    End If
End Sub
Public Function ReadMT940Statement(oFile As InputFile) As Boolean
    Dim sLine As String
    Dim sRest As String
    Dim stmt As New Statement
    Dim t As Txn
    Dim iTmp As Integer
    Dim sTmp As String
    Dim sLast As String
    Dim iStartLoc As Long
    Dim bInTxn As Boolean
    
'    On Error GoTo baderr
    
    Set t = New Txn
    Set t.Statement = stmt
    ReadMT940Statement = False
    If oFile.AtEOF Then Exit Function
    
    Do While Not oFile.AtEOF
        iStartLoc = oFile.Pos
        ShowProgress iStartLoc
        sLine = GetLine(oFile)
        If Left$(sLine, 1) = ":" Then
            iTmp = InStr(2, sLine, ":")
            sTmp = Left$(sLine, iTmp)
            sRest = Mid$(sLine, iTmp + 1)
            Select Case sTmp
            Case ":20:":
' start of new statement record
                If stmt.BankName <> "" Then
                    ' we have gone too far, back up a line
                    FlushTxn stmt, sTxn, sText, t
                    Set t = New Txn
                    Set t.Statement = stmt
                    oFile.Pos = iStartLoc
                    GoTo baleout
                End If
                stmt.BankName = Trim$(sRest)
                sLast = sTmp
            Case ":21:"
' related reference
                stmt.RelatedReference = sRest
                sLast = sTmp
            Case ":25:":
' find and validate account - format <swiftcode>/<account>
                iTmp = InStr(sRest, "/")
                If iTmp > 0 Then
                    stmt.BankName = Trim$(Left$(sRest, iTmp - 1))
                    stmt.Acct = Trim$(Mid$(sRest, iTmp + 1))
                Else
                    stmt.Acct = sRest
                End If
                sLast = sTmp
            Case ":28:", ":28C:":
                stmt.StatementID = sRest
                sTxn = ""
                sLast = sTmp
            Case ":60F:", ":60M:":
                If sText <> "" And sLast = ":NS:" Then
                    If bInTxn Then
                        t.NonSwift.ParseInfo sText
                    Else
                        stmt.NonSwift.ParseInfo sText
                    End If
                    sText = ""
                End If
' parse opening balance
                ParseBalance stmt.OpeningBalance, sRest
                sLast = sTmp
            Case ":61:":
                If sText <> "" And sLast = ":NS:" Then
                    If bInTxn Then
                        t.NonSwift.ParseInfo sText
                    Else
                        stmt.NonSwift.ParseInfo sText
                    End If
                    sText = ""
                End If
                FlushTxn stmt, sTxn, sText, t
                Set t = New Txn
                Set t.Statement = stmt
' collect new transaction
                sTxn = sRest: sText = ""
                bInTxn = True
                sLast = sTmp
            Case ":62F:", ":62M:":
                If sText <> "" And sLast = ":NS:" Then
                    If bInTxn Then
                        t.NonSwift.ParseInfo sText
                    Else
                        stmt.NonSwift.ParseInfo sText
                    End If
                    sText = ""
                End If
                FlushTxn stmt, sTxn, sText, t
                Set t = New Txn
                Set t.Statement = stmt
' insert closing balance
                ParseBalance stmt.ClosingBalance, sRest
' bale out now
                sLast = sTmp
            Case ":64:"
' available balance
                ParseBalance stmt.AvailableBalance, sRest
                sLast = sTmp
            Case ":65:"
' future available balance (can get several of these)
                sLast = sTmp
            Case ":86:"
                If sText = "" Then
                    sText = Trim$(sRest)
                Else
                    OtherLine sRest, sText
                End If
                sLast = sTmp
            Case ":NS:"
                sText = Trim$(sRest)
                sLast = sTmp
                Debug.Print ":NS: found:" & sRest
            Case Else
                Debug.Print "Unknown field """ & sTmp & """ found: " & sRest
                Debug.Assert False
                OtherLine sLine, sText
                ' ignore
            End Select
        ElseIf Left$(sLine, 1) = "-" Then
            If sLast = ":86:" Then  ' this is a comment line not end of stmt
                OtherLine sLine, sText
            Else
                FlushTxn stmt, sTxn, sText, t
                Set t = New Txn
' end of statement!
                Exit Do
            End If
        Else    ' does not start with ":"
            If sLast = ":86:" Then
                OtherLine sLine, sText
            ElseIf sLast = ":NS:" Then
                sText = sText & vbLf & sLine
' 20080401 CS: Capture "supplementary details" on line following :61:
' [x34] as per MT940 specs
            ElseIf sLast = ":61:" Then
                t.SupplementaryDetails = sLine
            ElseIf sLast = ":62F:" Or sLast = ":62M:" Then
                Exit Do
            End If
        End If
    Loop
baleout:
' 20050531 CS: transaction sorting now optional
    If Not Cfg.NoSortTxns Then
        InitProgress LoadResStringL(154), 0, 0
        stmt.Txns.SortByBookDate
    End If
    If sLast = "" Then
        ReadMT940Statement = False
    Else
        Session.Statements.Add stmt
        ReadMT940Statement = True
    End If
    Exit Function
baderr:
    MyMsgBox "An error has occurred"
End Function

Public Function ReadMT940File(oFile As InputFile) As Boolean
    InitProgress LoadResStringL(153), 0, oFile.Length
    sLastLine = ""
    sPrevLine = ""
    sLastColonLine = ""
    sPrevColonLine = ""
    iLineNum = 0
    While ReadMT940Statement(oFile)
'        myMsgBox "Statement read successfully"
    Wend
' if we baled out due to EOF all is well, otherwise an error has occurred
    ReadMT940File = oFile.AtEOF
End Function

Public Function GetLine(oFile As InputFile) As String
    Dim sLine As String
    sLine = oFile.ReadLine
    sPrevLine = sLastLine
    sLastLine = sLine
    If Left$(sLine, 1) = ":" Then
        sPrevColonLine = sLastColonLine
        sLastColonLine = sLine
    End If
    iLineNum = iLineNum + 1
    GetLine = sLine
End Function
