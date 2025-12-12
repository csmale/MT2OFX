' MT2OFX Input Processing Script Basic CSV format
' NB: This Script Will Not Work Without Customisation!

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/Base-CSV.vbs 10    20/04/08 10:01 Colin $"

Dim Params: Set Params = New MT2OFXScript

With Params
	.MinimumProgramVersion = "3"
	.DebugRecognition = False	' enables debug code in recognition
	.ScriptName = "Base-CSV"
	.FormatName = "Basic CSV"
	.ParseErrorMessage = "Cannot parse line."
	.ParseErrorTitle = .ScriptName
	.BankCode = ""
	.AccountNum = ""		' default if not specified in file
	.BranchCode = ""		' default if not specified in file
	.AccountType = "CREDITCARD"
	.QuickenBankID = ""		' copied to INTU.BID if present
	.CurrencyCode = "EUR"	' default if not specified in file
	.ColumnHeadersPresent = False	' are the column headers in the file?
	.SkipHeaderLines = 0	' number of lines to skip before the transaction data
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
	.CSVSeparator = ","
	.DecimalSeparator = ","	' as used in amounts
	.TxnLinePattern = ""
	.DateSequence = "DMY"	' must be DMY, MDY, or YMD
	.DateSeparator = "-/. "	' can be empty for dates in e.g. "yyyymmdd" format
	.OldestLast = True		' True if transactions are in reverse order
	.InvertSign = False	' make credits into debits etc
	.NoAvailableBalance = True		' True if file does not contain "Available Balance" information
	.MemoChunkLength = 0	' if memo field consists of fixed length chunks
	.TxnDatePattern = ".*(\d\d)\.(\d\d)\.(\d\d)\ (\d\d)\.(\d\d)"	' pattern to find transaction date in the memo
	.TxnDateSequence = Array(3,2,1,4,5,0)	' order of the info in the pattern (from 1 to 6): Y,M,D,H,M,S
	.PayeeLocation = 0		' start of payee in memo
	.PayeeLength = 0		' length of payee in memo
	.MonthNames = Empty
	.Fields = Array()
' min/max fields expected: default to size of Fields array. can be overridden here if required
'	.MinFieldsExpected = 1
'	.MaxFieldsExpected = 1
	.Properties = Empty
'	Set .TransactionCallback = GetRef("TransactionCallback")
'	Set .PreParseCallback = GetRef("PreParseCallback")
'	Set .HeaderCallback = GetRef("HeaderCallback")
'	Set .CustomDateCallback = GetRef("CustomDateCallback")
'	Set .CustomAmountCallback = GetRef("CustomAmountCallback")
'	Set .ReadLineCallback = GetRef("ReadLineCallback")
	Set .FinaliseCallback = GetRef("FinaliseCallback")
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

Private Function FindTag(xElement, sTag, sClass)
	Dim xTmp
	Dim xItem
	Set xTmp = xElement.getElementsByTagName(sTag)
	If Len(sClass) = 0 Then
		Set xItem = xTmp(0)
	Else
		For Each xItem In xTmp
			If xItem.ClassName = sClass Then
'				MsgBox "Found " & sTag & "/" & sClass
				Exit For
			End If
		Next
	End If
	Set FindTag = xItem
End Function

Private Function GetRow(xRow)
	Dim xCells, xCell
	Dim xArr(), i

	Set xCells = xRow.getElementsByTagName("TD")
	ReDim xArr(xCells.length - 1)
	i = 0
	For Each xCell In xCells
		xArr(i) = xCell.innerText
		i = i + 1
	Next
	GetRow = xArr
End Function

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Dim xDoc
Dim Stmt
Dim daDates: Set daDates = New DateAccumulator

Function RecogniseTextFile()
	RecogniseTextFile = False
	
	Set xDoc = GetHTMLDocument()
	If TypeName(xDoc) <> "HTMLDocument" Then
'		MsgBox "Got a " & TypeName(xDoc)
		Exit Function
	End If
	Dim xRekOverzicht
	Set xRekOverzicht = FindTag(xDoc, "DIV", "ICS_rekenoverzicht")
	If xRekOverzicht Is Nothing Then
		Exit Function
	End If
	
	RecogniseTextFile = True
	
	If RecogniseTextFile Then
		LogProgress ScriptName, "File Recognised"
	End If
End Function

Function LoadTextFile()
	LoadTextFile = False
	Dim vFields
	If xDoc Is Nothing Then
		If Not RecogniseTextFile() Then
			Exit Function
		End If
	End If
	Set Stmt = NewStatement()
	Stmt.OpeningBalance.Ccy = "EUR"
	Stmt.ClosingBalance.Ccy = "EUR"
	Stmt.BankName = Params.BankCode
	Stmt.BranchName = Params.BranchCode
	Stmt.AcctType = Params.AccountType

	Dim xRekOverzicht, xROHeader, xROBody, xRows, xRow
	Set xRekOverzicht = FindTag(xDoc, "DIV", "ICS_rekenoverzicht")
	If xRekOverzicht Is Nothing Then
		Exit Function
	End If
	Set xROHeader = FindTag(xRekOverzicht, "TABLE", "ICS_rekenoverzicht_header")
	Set xRows = xROHeader.getElementsByTagName("TR")

	Set xRow = xRows(2)
	vFields = GetRow(xRow)
	Stmt.Acct = vFields(2)
	
	Set xRow = xRows(5)
	vFields = GetRow(xRow)
	Dim sTmp, i
	For i = LBound(vfields) To UBound(vfields)
		sTmp = sTmp & vfields(i) & "-"
	Next
MsgBox "opening balance row: " & sTmp
	MsgBox "Opening Balance: " & vFields(1)
	sTmp = Replace(vFields(1), "debet", "")
	sTmp = Replace(sTmp, "credit", "")
	sTmp = Replace(sTmp, " ", "")
	Stmt.OpeningBalance.Amt = Params.ParseAmount(sTmp)
	If InStr(vFields(1), "debet") > 0 Then
		Stmt.OpeningBalance.Amt = -Stmt.OpeningBalance.Amt
	End If
	Set xRow = xRows(6)
	vFields = GetRow(xRow)
	For i = LBound(vfields) To UBound(vfields)
		sTmp = sTmp & vfields(i) & "-"
	Next
MsgBox "closing balance row: " & sTmp
	MsgBox "Closing Balance: " & vFields(3)
	sTmp = Replace(vFields(3), "debet", "")
	sTmp = Replace(sTmp, "credit", "")
	sTmp = Replace(sTmp, " ", "")
	Stmt.ClosingBalance.Amt = Params.ParseAmount(sTmp)
	If InStr(vFields(3), "debet") > 0 Then
		Stmt.ClosingBalance.Amt = -Stmt.ClosingBalance.Amt
	End If
	
	
	Set xROBody = FindTag(xRekOverzicht, "TABLE", "ICS_rekenoverzicht_body")
	Set xRows = xROBody.getElementsByTagName("TR")
	For Each xRow In xRows
		vFields = GetRow(xRow)
		If Not ProcessRow(vFields) Then
			Exit Function
		End If
	Next

	Stmt.OpeningBalance.BalDate = daDates.MinDate
	Stmt.ClosingBalance.BalDate = daDates.MaxDate

	LoadTextFile = True
End Function

Function ProcessRow(vFields)
	Dim dBookDate, dTxnDate, sMemo, dAmt
	Dim t
	ProcessRow = True
	
	Dim sTmp, i
	For i = LBound(vfields) To UBound(vfields)
		sTmp = sTmp & vfields(i) & "-"
	Next
'MsgBox "processing row: " & sTmp

	If UBound(vfields) < 5 Then
MsgBox "Not enough fields"
		Exit Function
	End If
	If vFields(1) = "" Then
MsgBox "empty first field"
		Exit Function
	End If

	dTxnDate = Params.ParseDate(vFields(1) & " 2008")
	If dTxnDate = NODATE Then
MsgBox "bad txn date " & vFields(1)
		Exit Function
	End If
	dTxnDate = MostRecent(dTxnDate)

	dBookDate = Params.ParseDate(vFields(2) & " 2008")
	If dBookDate = NODATE Then
MsgBox "bad book date " & vFields(2)
		Exit Function
	End If
	dBookDate = MostRecent(dBookDate)

	NewTransaction

	Txn.Payee = Trim(vFields(3))
	Txn.Memo = Txn.Payee
	ConcatMemo Trim(vFields(4))
	ConcatMemo CStr(dTxnDate)
	ConcatMemo Trim(vFields(5))
	dAmt = Params.ParseAmount(vFields(6))
	If Trim(vFields(7)) = "debet" Then
		dAmt = -dAmt
	End If	
	Txn.BookDate = dBookDate
	Txn.ValueDate = dBookDate
	Txn.TxnDate = dTxnDate
	Txn.TxnDateValid = True
	Txn.Amt = dAmt
	If Txn.Amt >= 0 Then
		Txn.TxnType = "DEP"
	Else
		Txn.TxnType = "PAYMENT"
	End If
	daDates.Process dBookDate
'	MsgBox "row: " & CStr(dBookDate) & " - " & sMemo & ": " & CStr(dAmt)
End Function

' callback functions, called from DefaultRecogniseTextFile and DefaultLoadTextFile
' ths following implementations are functionally neutral or equivalent to the default processing
' in the class
Function HeaderCallback(sLine)
MsgBox "In header callback: " & sLine
	HeaderCallback = True
End Function
Function TransactionCallback(t, vFields)
MsgBox "In transaction callback: " & t.Memo
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
Function FinaliseCallback()
MsgBox "In finalisation callback"
	FinaliseCallback = True
End Function
