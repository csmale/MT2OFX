Attribute VB_Name = "DoIt"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : DoIt
'    Project    : MT2OFX
'
'    Description: Main conversion processing flow
'
'    Modified   : $Author: Colin $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/DoIt.bas 30    15/11/10 0:27 Colin $"
' $History: DoIt.bas $
' 
' *****************  Version 30  *****************
' User: Colin        Date: 15/11/10   Time: 0:27
' Updated in $/MT2OFX
'
' *****************  Version 29  *****************
' User: Colin        Date: 6/10/09    Time: 0:33
' Updated in $/MT2OFX
' support for watcher - explicit output type
'
' *****************  Version 28  *****************
' User: Colin        Date: 15/06/09   Time: 19:24
' Updated in $/MT2OFX
' For transfer to new laptop
'
' *****************  Version 28  *****************
' User: Colin        Date: 17/01/09   Time: 23:24
' Updated in $/MT2OFX
' added option ForceRealBalanceDate
'
' *****************  Version 27  *****************
' User: Colin        Date: 25/11/08   Time: 22:17
' Updated in $/MT2OFX
' moving vss server!

'</CSCC>

' $Header: /MT2OFX/DoIt.bas 30    15/11/10 0:27 Colin $

Public Const ihMT940 As String = "builtin:MT940Input"
Public Const ohOFX As String = "builtin:OFXOutput"
Public Const ohQFX As String = "builtin:QFXOutput"
Public Const ohOFC As String = "builtin:OFCOutput"
Public Const ohQIF As String = "builtin:QIFOutput"
Public Const ohUnknown As String = ""

Public Function MT940ToOfx() As Boolean
    Dim t As Txn
    Dim iStmt As Integer
    Dim iTxn As Integer
    Dim iCount As Long
    Dim stmt As Statement
    Dim dBookDate As Date
    Dim dLastDate As Date
    Dim dEndDate As Date
    Dim iStmtSeq As Integer
    Dim sx As ScriptEnv
    Dim sTmp As String
    Dim sSetType As String
    Dim sCcy As String
    Dim dCBalDate As Date
    Dim dLastBalDate As Date

    MT940ToOfx = False
    
'    On Error GoTo baleout
    DBCSLog Session.FileOut, "Opening output OFX file"
    Session.OutputFile.CodePage = Cfg.OutputCodePage
    If Not Session.OutputFile.OpenFile(Session.FileOut) Then
        GoTo goback
    End If

    Set sx = GetScriptEnv()
    
    If Bcfg.BankKey <> "" And Bcfg.ScriptFile <> "" Then
        ScriptStartSession Session.BankID
        If sx.AbortRequested Then GoTo goback
    End If
    
    OFXSetVersion Cfg.OFXVersion
    
    OFXFileHeader Session.OutputFile, Session.ServerTime, Bcfg.IntuitBankID, _
        Session.Language
    
    sSetType = ""
    
    iCount = 0
    For iStmt = 1 To Session.Statements.Count
        If sx.AbortRequested Then GoTo goback
        Set stmt = Session.Statements(iStmt)
        dLastBalDate = NODATE
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

        If Bcfg.BankKey <> "" And Bcfg.ScriptFile <> "" Then
            ScriptProcessStatement Session.BankID, stmt
        End If

' keep sequence number for multiple statements on same date
        If dLastDate = stmt.ClosingBalance.BalDate Then
            iStmtSeq = iStmtSeq + 1
        Else
            iStmtSeq = 1
        End If
        dLastDate = stmt.ClosingBalance.BalDate

' get a unique ID for this statment. this should be in :28: but it is
' not reliable. if it is not usable, we create one based on year, and
' julian day.
        dEndDate = stmt.ClosingBalance.BalDate
        If stmt.OpeningBalance.BalDate > dEndDate Then
            dEndDate = stmt.OpeningBalance.BalDate
        End If
        With stmt
            If .OFXStatementID = "" Then
                If .StatementID = "" Or _
                    (.StatementNum = 0 And .StatementSeq = 1) Then
' CHANGE: the year gets prepended to OFXStatementID in FITID anyway so
' we don't need to put the year on here ourselves!
                    .OFXStatementID = Format(DatePart("y", .ClosingBalance.BalDate), "000") _
                    & Format(iStmtSeq, "00")
                Else
                    .OFXStatementID = .StatementID
                End If
            End If
        End With
        
        If stmt.BankName = "" Then
            stmt.BankName = "Unknown"
        End If
    ' make sure we have a currency for the statement header!
        If stmt.OpeningBalance.Ccy = "" Then
            stmt.OpeningBalance.Ccy = stmt.ClosingBalance.Ccy
        End If
        
' 20050904 CS: AcctType added to Statement, which (if non-empty) overrides the old logic
        If Len(stmt.AcctType) = 0 Then
            stmt.AcctType = AccountType(stmt.Acct)
        End If
        
' 20060322 CS: The transaction type now affects the file header (OFX message set declarator)
' 20060701 CS: handle message set changes properly!
        Debug.Assert Len(stmt.AcctType) > 0
        If sSetType <> stmt.AcctType Then
            If sSetType <> "" Then
                OFXStatementSetTrailer Session.OutputFile
            End If
            OFXStatementSetHeader Session.OutputFile, stmt.AcctType
            sSetType = stmt.AcctType
        End If
        
' 20041107 CS: Trim truncated bank name to avoid use of CDATA if it starts or ends
' with a space
' 20080610 CS: allow forcing of output currency in QFX (normally to USD)
        sCcy = stmt.OpeningBalance.Ccy
        If Session.FileFormat = FileFormatQFX And Len(Cfg.QFXForceCurrency) <> 0 Then sCcy = Cfg.QFXForceCurrency
        OFXStatementHeader Session.OutputFile, _
            sCcy, _
            Trim$(Left$(stmt.BankName, 9)), _
            Trim$(Left$(stmt.BranchName, 22)), _
            stmt.Acct, _
            stmt.AcctType, _
            stmt.OpeningBalance.BalDate, _
            dEndDate
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
' 20050918 CS: Added SIC
' 20051013 CS: Payee is now an object
' 20091203 CS: optionally output splits instead of main txn
            ' if mainonly then
            '   ofxtransaction
            ' elseif splitsonly then
            '   if splitspresent then
            '    ...
            '   else
            '    ofxtransaction
            '   endif
            ' elseif mainplussplits then
            '  if splitspresent then
            '   ofxtransaction (amt:=0)
            '   for each split
            '   next
            '  else
            '   ofxtransaction
            '  endif
            ' endif
            OFXTransaction Session.OutputFile, t.TxnType, _
                t.ValueDate, _
                t.TxnDate, _
                dBookDate, _
                t.TxnDateValid, _
                t.Amt, _
                t.FITID, _
                t.CheckNum, _
                t.Payee, _
                t.FurtherInfo, _
                t.SIC
            If dBookDate > dLastBalDate Then
                dLastBalDate = dBookDate
            End If
next_txn:
        Next
        dCBalDate = stmt.ClosingBalance.BalDate
        If stmt.ClosingBalance.Ccy <> "" Then
            If stmt.ClosingBalance.BalDate = NODATE Then
                If Cfg.OFXForceRealBalanceDate Then
                    dCBalDate = dLastBalDate
                End If
            End If
        End If
        OFXStatementTrailer Session.OutputFile, _
            stmt.ClosingBalance.Amt, _
            dCBalDate, _
            (stmt.AvailableBalance.Ccy <> ""), _
            stmt.AvailableBalance.Amt, _
            stmt.AvailableBalance.BalDate
skip_stmt:
    Next
    
    If Bcfg.BankKey <> "" And Bcfg.ScriptFile <> "" Then
        ScriptEndSession Session.BankID
    End If
    
    If sSetType <> "" Then
        OFXStatementSetTrailer Session.OutputFile
    End If
    
    OFXFileTrailer Session.OutputFile
    MT940ToOfx = True
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

Public Function MT940ToOfc() As Boolean
    Dim sCur As String
    Dim t As Txn
    Dim iStmt As Integer
    Dim iTxn As Integer
    Dim iCount As Long
    Dim stmt As Statement
    Dim dBookDate As Date
    Dim dLastDate As Date
    Dim iStmtSeq As Integer
    Dim sx As ScriptEnv
    Dim sTmp As String
    
    MT940ToOfc = False
    
'    On Error GoTo baleout
    DBCSLog Session.FileOut, "Opening output OFC file"
    Session.OutputFile.CodePage = Cfg.OutputCodePage
    If Not Session.OutputFile.OpenFile(Session.FileOut) Then
        GoTo goback
    End If

    Set sx = GetScriptEnv()
    
    If Bcfg.ScriptFile <> "" Then
        ScriptStartSession Session.BankID
        If sx.AbortRequested Then GoTo goback
    End If

    OFCFileHeader Session.OutputFile, Session.ServerTime
    
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
        End If

' keep sequence number for multiple statements on same date
        If dLastDate = stmt.ClosingBalance.BalDate Then
            iStmtSeq = iStmtSeq + 1
        Else
            iStmtSeq = 1
        End If
        dLastDate = stmt.ClosingBalance.BalDate

' get a unique ID for this statment. this should be in :28: but it is
' not reliable. if it is not usable, we create one based on year, and
' julian day.
        With stmt
            If .OFXStatementID = "" Then
                If .StatementID = "" Or _
                    (.StatementNum = 0 And .StatementSeq = 1) Then
' CHANGE: the year gets prepended to OFXStatementID in FITID anyway so
' we don't need to put the year on here ourselves!
                    .OFXStatementID = Format(DatePart("y", .ClosingBalance.BalDate), "000") _
                    & Format(iStmtSeq, "00")
                Else
                    .OFXStatementID = .StatementID
                End If
            End If
        End With
        
' 07 Nov 2004 Trim truncated bank name for symmetry with OFX
' 20041108 CS Added support for Branch ID
' 20060807 CS: AcctType added to Statement, which (if non-empty) overrides the old logic
        sTmp = OFCAccountType(stmt.Acct, stmt.AcctType)
        OFCStatementHeader Session.OutputFile, _
            sCur, _
            Trim$(Left$(stmt.BankName, 9)), _
            Trim$(Left$(stmt.BranchName, 9)), _
            stmt.Acct, _
            sTmp, _
            stmt.OpeningBalance.BalDate, _
            stmt.ClosingBalance.BalDate, _
            stmt.ClosingBalance.Amt
        For iTxn = 1 To stmt.Txns.Count
            If sx.AbortRequested Then GoTo goback
            Set t = stmt.Txns(iTxn)
            iCount = iCount + 1
            ShowProgress iCount
            
            If Not DoProcessTxn(Session.BankID, t) Then
                MyMsgBox GetString(117)
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
' 20050918 CS: added SIC
' 20051013 CS: Payee is now an object
            OFCTransaction Session.OutputFile, t.TxnType, _
                t.ValueDate, _
                t.TxnDate, _
                dBookDate, _
                t.TxnDateValid, _
                t.Amt, _
                t.FITID, _
                t.CheckNum, _
                t.Payee, _
                t.FurtherInfo, _
                t.SIC
next_txn:
        Next
        OFCStatementTrailer Session.OutputFile
skip_stmt:
    Next
    
    OFCFileTrailer Session.OutputFile
    If Bcfg.ScriptFile <> "" Then
        ScriptEndSession Session.BankID
    End If
    MT940ToOfc = True
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

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ProcessMt940
' Description:       Process an MT940 file
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       06/12/2003-21:27:38
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ProcessMt940() As Boolean
    Dim sBank As String
    ProcessMt940 = False
' work out which bank
    sBank = IdentifyBank(Session.InputFile)
    If sBank = "" Then
        Session.InputFile.CloseFile
        DoUnknownBank Session.FileIn
        Exit Function
    End If
    LogMessage False, True, "Bank: " & sBank, ""

' now we know the bank, get the bank-specific configuration
' into global Bcfg
    If Not (LoadBankSettings(sBank)) Then
        MyMsgBox GetString(116, sBank), vbCritical + vbOKOnly
        Session.InputFile.CloseFile
        Exit Function
    End If
    Session.BankID = sBank

    ProcessMt940 = ProcessAsMt940()
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ProcessAsMt940
' Description:       Do the actual MT940 processing - factored out for usage from scripts
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       4/19/2008-08:02:05
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ProcessAsMt940() As Boolean
' read the file!
    ProcessAsMt940 = False
    If Bcfg.BankKey = "" Then
        Bcfg.BankKey = Bcfg.IDString
    End If
    If Not ReadMT940File(Session.InputFile) Then
        MyMsgBox GetString(113, Session.FileIn), vbCritical + vbOKOnly
        Session.InputFile.CloseFile
        Exit Function
    End If
    ProcessAsMt940 = True
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ProcessText
' Description:       Process a text file
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       06/12/2003-21:27:04
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ProcessText(sScript As String) As Boolean
    Dim sMod As String
    ' set up dummy values in Bcfg
    Set Bcfg = New BankConfig
    Bcfg.ScriptFile = sScript
    
    Dim sx As ScriptEnv
    Set sx = GetScriptEnv()
    
    ' call the script function to do the work
    Session.InputFile.Rewind
    sMod = ScriptModuleName(Bcfg.ScriptFile)
    If Not InitialiseScripting() Then
        Session.InputFile.CloseFile
        Exit Function
    End If
    If Not ScriptInit(sMod, Bcfg.ScriptFile) Then
        LogMessage True, True, GetString(118), AppName
        Session.InputFile.CloseFile
        Exit Function
    End If
    
    LogMessage False, True, "Running script " & Bcfg.ScriptFile
    InitProgress LoadResStringL(150), 0, Session.InputFile.Length
    ProcessText = ScriptProcessTextFile(sMod)
    If sx.AbortRequested Then ProcessText = False
    If Not ProcessText Then
        sx.AbortRequested = True
        MyMsgBox GetString(113, Session.FileIn), vbCritical + vbOKOnly
        Session.InputFile.CloseFile
    End If
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Process
' Description:       Performs the actual conversion process
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       25/11/2003-21:22:46
'
' Parameters :       sFilein (String)
'                    sFileOut (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function Process(sFilein As String, sFileOut As String, iType As Long, sType As String) As Boolean
    ResetScripting
    Process = False
    Dim sx As ScriptEnv
    Set sx = GetScriptEnv()
    Set Session = New Session
    Set sx.Session = Session
    Session.FileIn = sFilein
    Session.FileOut = sFileOut
    Session.PayeeMapFile = Cfg.PayeeMapFile
    Session.PayeeMapIgnoreCase = Cfg.PayeeMapIgnoreCase
    Session.ServerTime = NODATE
    Session.Language = "ENG"

' 20050210 CS: Loading of payee map now in Session object so scripts
' can change the map file to be used dynamically
    If Session.PayeeMapFile <> "" Then
        LogMessage False, True, " IgnoreCase: " & CStr(Session.PayeeMapIgnoreCase)
    End If

    On Error GoTo process_err
    sx.AbortRequested = False
    DBCSLog sFilein, "Opening input file"
    Session.InputFile.OpenFile sFilein

    LogMessage False, True, "Input: " & sFilein, ""
    LogMessage False, True, "Output: " & sFileOut, ""
    
' prepare the output file type (file format)
    If Len(sType) = 0 And iType <> 5 Then
        Session.OutputFileType = GetDefaultOutputType(sFileOut)
    Else
        Session.CustomFormatName = sType
        Session.OutputFileType = sType
    End If
    LogMessage False, True, "Output File Type: " & Session.OutputFileType, ""

'   see if this file extension has been mapped explicitly to a script
' establish the input file type
    Dim iTmp As Long
    Dim sInpExt As String
    sInpExt = UCase$(GetExtension(sFilein))
    LogMessage False, True, "Input File Type: " & sInpExt, ""

    Dim sScript As String
    sScript = ScriptForInputFile()

'   if we have an input script, we treat it as a "text" file and load it
'   through that script.
'   otherwise it is assumed to be an MT940 file, which is loaded by the
'   standard mt940 routines and processed by the bank-specific scripts.

    If sScript = "" Then
        Process = ProcessMt940()
    Else
        LogMessage False, True, "Input code page: " & CStr(Session.InputFile.CodePage), ""
        Process = ProcessText(sScript)
    End If
    ' Bcfg must be set up by here!
    If Not Process Then GoTo wrapup
    
'======================================
'   OUTPUT PHASE STARTS HERE
'======================================

' load the processing scripts
    If Bcfg.ScriptFile <> "" And Bcfg.BankKey <> "" Then
        If Not InitialiseScripting() Then
            Session.InputFile.CloseFile
            CloseProgress
            Process = False
            Exit Function
        Else
            LogMessage False, True, "Loading script file: " & Bcfg.ScriptFile, ""
            If Not ScriptInit(Bcfg.BankKey, Bcfg.ScriptFile) Then
                LogMessage True, True, GetString(118), AppName
                Session.InputFile.CloseFile
                CloseProgress
                Process = False
                Exit Function
            End If
        End If
    End If

    If sx.AbortRequested Then
        Session.InputFile.CloseFile
        CloseProgress
        Exit Function
    End If
    
    Dim nTxns As Long
    Dim i As Long
    nTxns = 0
    With Session.Statements
        InitProgress LoadResStringL(152), 1, .Count
        For i = 1 To .Count
            nTxns = nTxns + .Item(i).Txns.Count
            ShowProgress i
        Next
    End With
    InitProgress LoadResStringL(151), 0, nTxns
    
' centralise session initialisation
    If Session.ServerTime = NODATE Then
        Session.ServerTime = FindServerTime(Session.FileIn)
    End If

' 20050220 CS: Add language code to session
    If Len(Session.Language) <> 3 Then
        Session.Language = "ENG"
    End If
    
    Select Case Session.FileFormat
    Case FileFormatOFX, FileFormatQFX
        Process = MT940ToOfx()
    Case FileFormatOFC
        Process = MT940ToOfc()
    Case FileFormatQIF
        Process = MT940ToQIF()
    Case FileFormatCustom
        Process = MT940ToCustom(sType)
    Case Else
        LogMessage False, True, "Unimplemented output type '" & Session.OutputFile & "'"
        Process = False
    End Select
wrapup:
    CloseProgress
    Session.InputFile.CloseFile
    If Process Then
        LogMessage False, True, "Read " & Session.Statements.Count & " statements"
        LogMessage False, True, "Read " & Session.TransactionCount & " transactions"
        LogMessage False, True, "Total credits: " & Session.TotalCredits
        LogMessage False, True, "Total debits: " & Session.TotalDebits
        If Session.SuppressedStatementCount > 0 Then
            LogMessage False, True, "Skipped " & Session.SuppressedStatementCount & " empty statements"
        End If
    End If
    Exit Function
process_err:
    ShowError "Process"
    Process = False
'    MyMsgBox Err.Description
    Resume wrapup
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoUnknownBank
' Description:       Handle the fact that the file cannot be recognised
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       06/12/2003-21:20:52
'
' Parameters :       sFileIn (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function DoUnknownBank(sFilein As String) As Boolean
    Dim f As New frmUnkBank
    f.FileName = sFilein
    f.Show vbModal
    Set f = Nothing
    DoUnknownBank = True
End Function

Public Function NewProcess(sFilein As String, sFileOut As String) As Boolean

' input file has been selected in sFileIn
' output file has been selected in sFileOut

' initialisesession
    ResetScripting
    NewProcess = False
    Dim sx As ScriptEnv
    Set sx = GetScriptEnv()
    Set Session = New Session
    Set sx.Session = Session
    Session.FileIn = sFilein
    Session.FileOut = sFileOut
    Session.PayeeMapFile = Cfg.PayeeMapFile
    Session.PayeeMapIgnoreCase = Cfg.PayeeMapIgnoreCase
    Session.ServerTime = NODATE
    Session.Language = "ENG"
' 20050210 CS: Loading of payee map now in Session object so scripts
' can change the map file to be used dynamically
    If Session.PayeeMapFile <> "" Then
        LogMessage False, True, " IgnoreCase: " & CStr(Session.PayeeMapIgnoreCase)
    End If

'    On Error GoTo process_err
    sx.AbortRequested = False

' identify input handler
    DBCSLog sFilein, "Opening input file"
    Session.InputFile.OpenFile sFilein
    LogMessage False, True, "Input: " & sFilein, ""

' identify output handler
    LogMessage False, True, "Output: " & sFileOut, ""


' read file using InputHandler
' process file
'  mt940 scripting using ProcessHandler?
'  payee replacement
' write output file using OutputHandler

End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       MT940ToCustom
' Description:       Process to custom output format
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       08/05/2010-18:45:15
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function MT940ToCustom(sType As String) As Boolean
    Dim sx As ScriptEnv
    Dim sMod As String
    Dim iCount As Long
    Dim iStmt As Long
    Dim stmt As Statement
    Dim dLastDate As Date
    Dim dLastBalDate As Date
    Dim iStmtSeq As Long
    Dim dEndDate As Date
    Dim iTxn As Long
    Dim t As Txn
    Dim dBookDate As Date
    Dim dCBalDate As Date
    Dim sSetType As String
    
    MT940ToCustom = True

    Set sx = GetScriptEnv()
    
' only basic stuff
' initialise script as per sType
' then call script
    
    sMod = ScriptModuleName(sType)
    If Not InitialiseScripting() Then
        Session.InputFile.CloseFile
        Exit Function
    End If
    If Not ScriptInit(sMod, sType) Then
        LogMessage True, True, GetString(118), AppName
        Session.InputFile.CloseFile
        Exit Function
    End If
    LogMessage False, True, "Using custom output script " & sType
    
    Session.OutputFile.CodePage = Cfg.OutputCodePage
    If Not Session.OutputFile.OpenFile(Session.FileOut) Then
        GoTo goback
    End If

    MT940ToCustom = ScriptCustomOutput(sMod, Session)
    
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
