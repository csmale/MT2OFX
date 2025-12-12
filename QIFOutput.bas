Attribute VB_Name = "QIFOutput"
Option Explicit

' $Header: /MT2OFX/QIFOutput.bas 22    15/11/10 0:13 Colin $

Private sDateFmt As String
Private iGMTOffset As Long  ' in minutes

Private Function FormatQIFDate(dDate As Date) As String
    Dim dTmp As Date
    dTmp = dDate
' if there is a non-zero time part, this time is in GMT and needs to be corrected to the local timezone
' in case this causes the date to change to the next/previous day!
    If (Hour(dTmp) + Minute(dTmp) + Second(dTmp)) > 0 Then
        dTmp = DateAdd("n", -iGMTOffset, dTmp)
    End If
    FormatQIFDate = Format(dTmp, sDateFmt)
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       MT940ToQIF
' Description:       Save output in QIF format
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       22/10/2003-22:59:34
'
' Parameters :       none
'--------------------------------------------------------------------------------
'</CSCM>
Public Function MT940ToQIF() As Boolean
    Dim X As Variant
    Dim sCur As String
    Dim t As Txn
    Dim iStmt As Integer
    Dim iTxn As Integer
    Dim stmt As Statement
    Dim iFileOut As Integer
    Dim dBookDate As Date
    Dim dLastDate As Date
    Dim iStmtSeq As Integer
    Dim sLastAcct As String
    Dim sAcctType As String
    Dim sTmp As String
    Dim iDateFmt As Integer
'    Dim sDateFmt As String
    Dim sDecSep As String
    Dim sSysDecSep As String
    Dim i As Long
    Dim iCount As Long
    Dim sx As ScriptEnv
    Set sx = GetScriptEnv()
    
    MT940ToQIF = False

    If Cfg.QifDateFormat = DATEFMT_SYSTEM Then
        iDateFmt = SystemShortDateFormat()
    Else
        iDateFmt = Cfg.QifDateFormat
    End If
    Select Case iDateFmt
    Case DATEFMT_DMY
        sDateFmt = "dd/mm/yyyy"
    Case DATEFMT_MDY
        sDateFmt = "mm/dd/yyyy"
    Case DATEFMT_YMD
        sDateFmt = "yyyy/mm/dd"
    Case DATEFMT_CUSTOM
        sDateFmt = Cfg.QifCustomDateFormat
    Case Else
        Debug.Assert False
    End Select
    
    sSysDecSep = SystemDecimalSeparator()
    sDecSep = Cfg.QifDecimalSeparator
    If sDecSep = "" Then
        sDecSep = sSysDecSep
    End If
    iGMTOffset = GetCurrentTimeBias()   ' returns minutes from GMT
    
    On Error GoTo baleout
    DBCSLog Session.FileOut, "Opening QIF output file"
    If Not Session.OutputFile.OpenFile(Session.FileOut) Then
        GoTo goback
    End If
    Session.OutputFile.CodePage = Cfg.OutputCodePage
    
'    iFileOut = FreeFile
'    Open Session.FileOut For Output Access Write As iFileOut
    
    If Bcfg.ScriptFile <> "" Then
        ScriptStartSession Session.BankID
        If sx.AbortRequested Then GoTo goback
    End If
    
' 20070210 CS: Initialise account type with some junk value to make absolutely sure that it gets a value,
' even if the account number is empty
    sLastAcct = "famous_last_words"
    iCount = 0
    For iStmt = 1 To Session.Statements.Count
        If sx.AbortRequested Then GoTo goback
        Set stmt = Session.Statements(iStmt)

' 20060825 CS: added skipping empty statements (except last one for each account?)
        If Cfg.SuppressEmptyStatements Then
            If stmt.Txns.Count = 0 Then
                If iStmt = Session.Statements.Count Then
                    GoTo no_skip
                End If
                If stmt.Acct <> Session.Statements(iStmt + 1).Acct Then
                    GoTo no_skip
                End If
                Session.SuppressedStatementCount = Session.SuppressedStatementCount + 1
                GoTo skip_stmt
            End If
        End If
no_skip:

        sCur = stmt.OpeningBalance.Ccy

        If Bcfg.ScriptFile <> "" Then
            ScriptProcessStatement Session.BankID, stmt
            If sx.AbortRequested Then GoTo goback
        End If
        
        If stmt.Acct <> sLastAcct Then
            If Len(stmt.AcctType) = 0 Then
                stmt.AcctType = AccountType(stmt.Acct)
            End If
            Select Case stmt.AcctType
            Case "CHECKING"
                sAcctType = "Bank"
            Case "SAVINGS"
                sAcctType = "Bank"
            Case "CREDITLINE"
                sAcctType = "Oth L"
            Case "MONEYMRKT"
                sAcctType = "Oth A"
            Case "CMA"
                sAcctType = "Oth A"
            Case "CREDITCARD"   ' special for MT2OFX - not OFX standard
                sAcctType = "CCard"
            Case Else
                sAcctType = "Bank"
            End Select
' 20051014 CS: Added QIFAcctType override so the script can determint
            If stmt.QIFAcctType <> "" Then sAcctType = stmt.QIFAcctType
            If Not Cfg.QifNoAcctHeader Then
                Session.OutputFile.PrintLine "!Account"
                Session.OutputFile.PrintLine "N" & stmt.Acct
                Session.OutputFile.PrintLine "T" & sAcctType
'            Print #iFileOut, "DOFX online account"
' 7 Oct 2004: if no closing balance is known, don't output misleading rubbish
                If stmt.ClosingBalance.BalDate > DateSerial(1950, 0, 0) Then
                    Session.OutputFile.PrintLine "/" & FormatQIFDate(stmt.ClosingBalance.BalDate)
                    Session.OutputFile.PrintLine "$" & FormatAmount(stmt.ClosingBalance.Amt, sDecSep)
                End If
                Session.OutputFile.PrintLine "^"
            End If
            sLastAcct = stmt.Acct
        End If

        If stmt.Txns.Count > 0 Then
            Session.OutputFile.PrintLine "!Type:" & sAcctType
        End If
        For iTxn = 1 To stmt.Txns.Count
            If sx.AbortRequested Then GoTo goback
            Set t = stmt.Txns(iTxn)
            
            iCount = iCount + 1
            ShowProgress iCount
            
            If Not DoProcessTxn(Session.BankID, t) Then
'                MsgBox GetString(117)
                GoTo next_txn
            End If
            
            Select Case Cfg.BookDateMode
            Case bdmBookDate
                dBookDate = t.BookDate
            Case bdmValueDate
                dBookDate = t.ValueDate
            Case bdmTransDate
                If t.TxnDateValid Then
                    dBookDate = t.TxnDate
                Else
                    dBookDate = t.BookDate
                End If
            Case Else
                Debug.Assert False
                dBookDate = t.BookDate
            End Select
' 7 Oct 2004: Sequence of elements modified to mimic MS Money (for ease of debugging)
            Session.OutputFile.PrintLine "D" & FormatQIFDate(dBookDate)
' 12 August 2004 - added ClearedStatus for use with QIF
            If Len(t.ClearedStatus) > 0 Then
                Session.OutputFile.PrintLine "C" & t.ClearedStatus
            End If
' 7 Oct 2004 - don't output an empty M line
            If Len(t.Memo) > 0 Then
' 20050126 CS: Added QifMaxMemoLength
                If Cfg.QifMaxMemoLength > 0 Then
                    Session.OutputFile.PrintLine "M" & Left$(t.Memo, Cfg.QifMaxMemoLength)
                Else
                    Session.OutputFile.PrintLine "M" & t.Memo
                End If
            End If
            Session.OutputFile.PrintLine "T" & FormatAmount(t.Amt, sDecSep)
            If Cfg.QifOutputUAmount Then
                Session.OutputFile.PrintLine "U" & FormatAmount(t.Amt, sDecSep)
            End If
            If t.CheckNum <> "" Then
                Session.OutputFile.PrintLine "N" & t.CheckNum
            ElseIf t.BankReference <> "" Then
                Session.OutputFile.PrintLine "N" & t.BankReference
            End If
' CS 20050109: Make Payee optional
            If Len(t.Payee) > 0 Then
                Session.OutputFile.PrintLine "P" & t.Payee
                For i = 1 To t.Payee.LastUsedAddressLine
                    Session.OutputFile.PrintLine "A" & t.Payee.Addr(i)
                Next
            End If
' 11 May 2004 - added Category for use with QIF
            If Len(t.Category) > 0 Then
                Session.OutputFile.PrintLine "L" & t.Category
            End If
'            Session.OutputFile.PrintLine "L" & "unknown transaction type"
' 27 Oct 2004 - added Split output
            Dim s As Split
            For Each s In t.Splits
                If Len(s.Category) > 0 Then
                    Session.OutputFile.PrintLine "S" & s.Category
                End If
                If Len(s.Memo) > 0 Then
' 20050126 CS: Added QifMaxMemoLength
                    If Cfg.QifMaxMemoLength > 0 Then
                        Session.OutputFile.PrintLine "E" & Left$(s.Memo, Cfg.QifMaxMemoLength)
                    Else
                        Session.OutputFile.PrintLine "E" & s.Memo
                    End If
                End If
                Session.OutputFile.PrintLine "$" & FormatAmount(s.Amt, sDecSep)
            Next
' 20051013 CS: Added CustomQIFItems
            For Each X In t.CustomQIFItems
                Session.OutputFile.PrintLine CStr(X) & t.CustomQIFItems(X)
            Next
            Session.OutputFile.PrintLine "^"
next_txn:
        Next
skip_stmt:
    Next
    
    If Bcfg.ScriptFile <> "" Then
        ScriptEndSession Session.BankID
        If sx.AbortRequested Then GoTo goback
    End If

    MT940ToQIF = True
goback:
    Session.OutputFile.CloseFile
    Exit Function
baleout:
    MyMsgBox Err.Description
    Resume goback
open_err:
    MyMsgBox GetString(114, Session.FileIn, Err.Description)
    Resume goback
End Function

