' MT2OFX Input Processing Script AAB CAMT.053(ZIP) format

Option Explicit

Const ScriptVersion = "$Header$"

Dim Params: Set Params = New MT2OFXScript

With Params
	.MinimumProgramVersion = "3"
	.DebugRecognition = False	' enables debug code in recognition
	.ScriptName = "CAMT053-XML-ZIP"
	.FormatName = "CAMT.053"
	.ParseErrorMessage = "Cannot parse line."
	.ParseErrorTitle = .ScriptName
   .CodePage = 1252  ' Windows English / Western Europe
	.BankCode = "ABNANL2A"
	.AccountNum = ""		' default if not specified in file
	.BranchCode = ""		' default if not specified in file
	.AccountType = "CHECKING"	' an be CHECKING or CREDITCARD
	.QuickenBankID = ""		' copied to INTU.BID if present
	.CurrencyCode = "EUR"	' default if not specified in file
	.ColumnHeadersPresent = True	' are the column headers in the file?
	.SkipHeaderLines = 0	' number of lines to skip before the transaction data
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
	.CSVSeparator = ","
	.DecimalSeparator = "."	' as used in amounts
	.TxnLinePattern = ""
	.DateSequence = "YMD"	' must be DMY, MDY, or YMD
	.DateSeparator = "-/. "	' can be empty for dates in e.g. "yyyymmdd" format
	.OldestLast = True		' True if transactions are in reverse order
	.InvertSign = False	' make credits into debits etc
	.NoAvailableBalance = False		' True if file does not contain "Available Balance" information
	.MemoChunkLength = 0	' if memo field consists of fixed length chunks
	.TxnDatePattern = ".*(\d\d)\.(\d\d)\.(\d\d)\ (\d\d)\.(\d\d)"	' pattern to find transaction date in the memo
	.TxnDateSequence = Array(3,2,1,4,5,0)	' order of the info in the pattern (from 1 to 6): Y,M,D,H,M,S
	.PayeeLocation = 0		' start of payee in memo
	.PayeeLength = 0		' length of payee in memo
	.MonthNames = Empty
' Date,Transaction Type,Check Number,Description,Amount
	.Fields = Array( _
		Array(fldBookDate, "Date"), _
		Array(fldSkip, "Transaction Type"), _
		Array(fldCheckNum, "Check Number"), _
		Array(fldMemo, "Description"), _
		Array(fldAmount, "Amount") _
	)
' min/max fields expected: default to size of Fields array. can be overridden here if required
'	.MinFieldsExpected = 1
'	.MaxFieldsExpected = 1
	.Properties = Array( _
		Array("AcctNum", "Account number", _
			"The account number for " & Params.FormatName, _
			ptString,,"=CheckAccount", "Please enter a valid account number.") _
		)
'  Set .TransactionCallback = GetRef("TransactionCallback")
'	Set .IsValidTxnLine = GetRef("IsValidTxnLine")
'	Set .PreParseCallback = GetRef("PreParseCallback")
'	Set .HeaderCallback = GetRef("HeaderCallback")
'	Set .StatementCallback = GetRef("StatementCallback")
'	Set .CustomDateCallback = GetRef("CustomDateCallback")
'	Set .CustomAmountCallback = GetRef("CustomAmountCallback")
'	Set .ReadLineCallback = GetRef("ReadLineCallback")
End With

'MonthNames = Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
' Either give the month names in an array as above or use SetLocale to get the
' system strings for the given locale. Otherwise the default locale will be used.
' The MonthNames array must have a multiple of 12 elements, which run from Jan-Dec in groups of
' 12, i.e. "Jan".."Dec","January".."December" etc. Lower/upper case is not significant.
' SetLocale "nl-nl"

' Field name constants are now in MT2OFX.vbs
' For reference, they are:
' fldSkip, fldAccountNum, fldCurrency, fldClosingBal, fldAvailBal,
' fldBookDate, fldValueDate, fldAmtCredit, fldAmtDebit, fldMemo
' fldBalanceDate, fldAmount, fldPayee, fldTransactionDate, fldTransactionTime,
' fldChequeNum, fldCheckNum, fldFITID, fldEmpty, fldBranch

' Declare fields in the order they appear in the file as an array of arrays. The inner arrays
' contain a field ID from the list above followed by the exact column header.
' An optional third element in the inner arrays is used to contain a RegExp pattern for use instead of the
' literal text in the second element. If the pattern starts with "=", it is treated as a VBScript expression,
' where the characters "%1" are replaced with the contents of the field from the file.
' For example: "=Validate(""%1"")" would cause the function Validate to be called, which must return either
' True or False to indicate whether the validation passed.


' Property List is an array of arrays, each of which has the following elements:
'	1. Property key - used to reference properties
'	2. Property name - used as a label in the config screen
'	3. Property description - used as a description or tooltip in the config screen
'	4. Data type - ptString, ptBoolean, ptInteger, ptFloat, ptDate, ptChoose
'	5. Value list (will be displayed in a combobox) - array of values (Only with ptChoose)
'	6. Validation pattern (optional) - RegExp to validate the value entered
'		If the pattern starts with "=", the rest of the string is taken to be the name of a function in this
'		script which is called, with the value entered as a parameter, and which must return True if the value
'		is acceptable and False otherwise.
'	7. Validation error message (optional) - Message which will be displayed if the value entered fails the validation.
'		The script may instead define a function ValidationMessage which must return a string containing the message.
'		In both cases, "%1" in the string will be replaced by the value entered.

Dim oShell
Dim sTmpDir
Dim fso

Function CheckAccount(s)
	CheckAccount = False
	If Len(s) = 0 Then Exit Function
	CheckAccount = True
End Function

Sub Initialise()
    LogProgress ScriptVersion, "Initialise"
    If Not CheckVersion() Then
        Abort
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    sTmpDir = fso.GetSpecialFolder(2) ' gets temp directory from system  
    If Not fso.FolderExists(sTmpDir) Then
        fso.CreateFolder(sTmpDir)
    End If
End Sub

' function DescriptiveName
' returns a string with a descriptive name of this script
Function DescriptiveName()
	DescriptiveName = FormatName
End Function

Sub Configure
	If ShowConfigDialog(Params.ScriptName, Params.Properties) Then
		SaveProperties Params.ScriptName, Params.Properties
	End If
End Sub

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
    RecogniseTextFile = False
    If LCase(Right(Session.FileIn,4)) <> ".zip" Then
        Exit Function
    End If
    ' open the zip file using Shell Namespace object
    Dim oShell: Set oShell = CreateObject("Shell.Application")
    Dim oZip: Set oZip = oShell.NameSpace(Session.FileIn).Items
    If oZip.Count < 1 Then
        MsgBox "Empty ZIP file"
        Exit Function
    End If
    Dim sXML
    sXML = Mid(oZip.Item(0).Path, Len(Session.FileIn)+2)
    Dim re: Set re = New RegExp
    re.Pattern = "\d{8}_\d{9}_\d{12}.xml"
    If Not re.Test(sXML) Then
        MsgBox "Unexpected file name " & sXML
        Exit Function
    End If
    
    'MsgBox "Full Path=" & oZip.Item(0).Path
    
    Dim xDoc: Set xDoc = CreateObject("MSXML2.DOMDocument")
    xDoc.async = False
    Dim oTmpDir: Set oTmpDir = oShell.NameSpace(sTmpDir)
    oTmpDir.CopyHere oZip.Item(0), 4+16
    Dim sTmpFile: sTmpFile = sTmpDir & "\" & Mid(oZip.Item(0).Path,  Len(Session.FileIn)+2)
    If Not xDoc.Load(sTmpFile) Then
        If xDoc.parseError.errorCode <> 0 Then
            Dim myErr: Set myErr = xDoc.parseError
            MsgBox "You have error " + myErr.reason
        Else
            MsgBox xDoc.parseError.reason
        End If
        MsgBox "Unable to load XML from " & sTmpFile
        fso.DeleteFile sTmpFile, True
        Exit Function
    End If
    xDoc.setProperty "SelectionLanguage", "XPath"
    xDoc.setProperty "SelectionNamespaces", "xmlns:s='urn:iso:std:iso:20022:tech:xsd:camt.053.001.02'"
    If xDoc.SelectNodes("/s:Document/s:BkToCstmrStmt").length < 1 Then
        MsgBox "Content is not CAMT.053"
        fso.DeleteFile sTmpFile, True
        Exit Function
    End If
    
'    MsgBox "opened XML"
    RecogniseTextFile = True
    
    fso.DeleteFile sTmpFile, True
    If RecogniseTextFile Then
        LogProgress ScriptName, "File Recognised"
    End If
End Function

Function LoadTextFile()
    LoadTextFile = False
    ' open the zip file using Shell Namespace object
    Set oShell = CreateObject("Shell.Application")
    Dim oZip: Set oZip = oShell.NameSpace(Session.FileIn).Items
    Dim re: Set re = New RegExp
    re.Pattern = "\d{8}_\d{9}_\d{12}.xml" ' this is for abn amro, dont know about others
    Dim i, oItem, sXML
    For i = 0 To oZip.Count-1
        Set oItem = oZip.Item(i)
        sXML = oItem.Path
        If (Not oItem.IsFolder) And re.Test(sXML) Then
        ' do this one!
            If Not CAMT053File(oItem) Then
                Exit For
            End If
        End If
    Next
    LoadTextFile = True
End Function

Function CAMT053File(oItem)
    Dim Stmt
    Dim oTmpDir: Set oTmpDir = oShell.NameSpace(sTmpDir)
    oTmpDir.CopyHere oItem, 4+16
    Dim sTmpFile: sTmpFile = sTmpDir & "\" & Mid(oItem.Path,  Len(Session.FileIn)+2)
    Dim xDoc: Set xDoc = CreateObject("MSXML2.DOMDocument")
    xDoc.async = False
    If Not xDoc.Load(sTmpFile) Then
        If xDoc.parseError.errorCode <> 0 Then
            Dim myErr: Set myErr = xDoc.parseError
            MsgBox "XML parse error " + myErr.reason
        Else
            MsgBox xDoc.parseError.reason
        End If
        MsgBox "Unable to load XML from " & sTmpFile
        fso.DeleteFile sTmpFile, True
        Exit Function
    End If
    xDoc.setProperty "SelectionLanguage", "XPath"
    xDoc.setProperty "SelectionNamespaces", "xmlns:s='urn:iso:std:iso:20022:tech:xsd:camt.053.001.02'"
    
    Dim xMsg, xStmt, xBal, xTxn, sCcy, sBalType, sTmp
    For Each xMsg In xDoc.SelectNodes("/s:Document/s:BkToCstmrStmt")
' group header here
        sTmp = NodeText(xMsg.SelectSingleNode("s:GrpHdr/s:CreDtTm"))
        If Len(sTmp) > 0 Then
            Session.ServerTime = ParseISO8601(sTmp)
        End If
        For Each xStmt In xMsg.SelectNodes("s:Stmt")
            Set Stmt = NewStatement()
            Stmt.Acct = NodeText(xStmt.SelectSingleNode("s:Acct/s:Id/s:IBAN"))
            Stmt.BankName = NodeText(xStmt.SelectSingleNode("s:Acct/s:Svcr/s:FinInstnId/s:BIC"))
            Stmt.StatementID = NodeText(xStmt.SelectSingleNode("s:Id"))
            sCcy = NodeText(xStmt.SelectSingleNode("s:Acct/s:Ccy"))
            For Each xBal In xStmt.SelectNodes("s:Bal")
                sBalType = NodeText(xBal.SelectSingleNode("s:Tp/s:CdOrPrtry/s:Cd"))
                Select Case sBalType
                Case "CLBD"   ' closing booked
                    ParseBalance xBal, Stmt.ClosingBalance
                Case "CLAV"   ' closing available
                    ParseBalance xBal, Stmt.AvailableBalance
                Case "OPBD", "PRCD" ' opening booked
                    ParseBalance xBal, Stmt.OpeningBalance
                End Select
            Next
            For Each xTxn In xStmt.SelectNodes("s:Ntry")
                If NodeText(xTxn.SelectSingleNode("s:Sts")) = "BOOK" Then
                    NewTransaction
                    Txn.Amt = GetAmount(xTxn, sCcy)
                    Txn.BookDate = GetDateTime(xTxn.SelectSingleNode("s:BookgDt"))
                    Txn.ValueDate = GetDateTime(xTxn.SelectSingleNode("s:ValDt"))
                    Txn.FITID = NodeText(xTxn.SelectSingleNode("s:AcctSvcrRef"))
                    Txn.Payee = NodeText(xTxn.selectSingleNode("s:NtryDtls/s:TxDtls/s:RltdPties/s:Cdtr/s:Nm"))
                    Txn.Memo = NodeText(xTxn.SelectSingleNode("s:AddtlNtryInf"))
                End If
            Next
        Next
    Next
    fso.DeleteFile sTmpFile, True
    CAMT053File = True
End Function

Function GetAmount(x, sCcy)
    Dim amt
    amt = Params.ParseAmount(NodeText(x.SelectSingleNode("s:Amt")))
    If NodeText(x.SelectSingleNode("s:CdtDbtInd")) = "DBIT" Then
        amt = -amt
    End If
    sCcy = x.SelectSingleNode("s:Amt").GetAttribute("Ccy")
    GetAmount = amt
End Function

Function GetDateTime(x)
    Dim sDt
    sDt = NodeText(x.SelectSingleNode("s:Dt"))
    If Len(sDt) = 0 Then
        sDt = NodeText(xBal.SelectSingleNode("s:DtTm"))
        GetDateTime = ParseISO8601(sDt)
    Else
        GetDateTime = Params.ParseDate(sDt)
    End If
End Function

Function ParseISO8601(sDt)
    Dim d, sTz
    d = DateSerial(Left(sDt, 4), Mid(sDt, 6, 2), Mid(sDt, 9, 2))
    If Len(sDt) > 10 Then
        d = d + TimeSerial(Mid(sDt, 12, 2), Mid(sDt, 15, 2), Mid(sDt, 18, 2))
        If Len(sDt) > 19 Then
'            sTz = Mid(sDt, 20)
        End If
    End If
    ParseISO8601 = d
End Function

Sub ParseBalance(xBal, b)
    b.Amt = GetAmount(xBal, b.Ccy)
    b.BalDate = GetDateTime(xBal.SelectSingleNode("s:Dt"))
End Sub

Function GetUTCOffset() 
    Dim oWSH : Set oWSH = CreateObject("WScript.Shell") 
    GetUTCOffset = oWSH.RegRead("HKLM\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias") 
End Function 

Function NodeText(x)
    If x Is Nothing Then
        NodeText = ""
    Else
        NodeText = x.Text
    End If
End Function

Function ParseSepaInfo(sLine)
    Dim aDict: Set aDict = CreateObject("Scripting.Dictionary")
    Dim aVal: aVal = Split(sLine, "/")
    Dim i, sKey, sVal
    For i=0 To UBound(aVal) Step 2
        sKey = aVal(i)
        sVal = aVal(i+1)
        aDict(sKey) = sVal
    Next
    Set ParseSepaInfo = aDict
End Function

' callback functions, called from DefaultRecogniseTextFile and DefaultLoadTextFile
' ths following implementations are functionally neutral or equivalent to the default processing
' in the class
Function HeaderCallback(sLine)
MsgBox "In header callback: " & sLine
	HeaderCallback = True
End Function
Function TransactionCallback(t, vFields)
'MsgBox "In transaction callback: " & t.Memo
    TransactionCallback = True
End Function
Function CustomDateCallback(sDate)
MsgBox "In custom date callback: " & sDate
	CustomDateCallback = ParseDateEx(sDate, Params.DateSequence, Params.DateSeparator)
End Function
Function CustomAmountCallback(sAmt)
MsgBox "In custom amount callback: " & sAmt
	CustomAmountCallback = ParseNumber(sAmt, Params.DecimalSeparator)
End Function
Function ReadLineCallback(sLine)
MsgBox "In read line callback: " & sLine
	ReadLineCallback = sLine
End Function
' PreParseCallback: returns True or False. True means the line can be processed; False means skip this line.
Function PreParseCallback(vFields)
MsgBox "In preparse callback: " & UBound(vFields) & " fields."
	PreParseCallback = True
End Function
Sub StatementCallback(Stmt)
MsgBox "In statement callback"
End Sub
Function FinaliseCallback()
MsgBox "In finalisation callback"
	FinaliseCallback = True
End Function
' IsValidTxnLine can return:
' txnlineSKIP: skip this line
' txnlineNORMAL: this is the first line of a new transaction
' txnlineCONTINUATION: this line continues the previous transaction
' if using continuation lines, set Params.Fields and Params.TxnLinePattern before returning!
Function IsValidTxnLine(sLine)
MsgBox "In IsValidTxnLine callback: " & sLine
	IsValidTxnLine = Iif(IsNumeric(Left(sLine, 1)), txnlineNORMAL, txnlineSKIP)
End Function
