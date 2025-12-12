' MT2OFX Input Processing Script Basic CSV format
' NB: This Script Will Not Work Without Customisation!

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/ICSVisaNL-PDFClip.vbs 1     12/01/09 22:38 Colin $"

Dim Params: Set Params = New MT2OFXScript

With Params
	.MinimumProgramVersion = "3"
	.DebugRecognition = False	' enables debug code in recognition
	.ScriptName = "ICSVisaNL-PDFClip"
	.FormatName = "Visa (ICS) copy text from PDF to clipboard"
	.ParseErrorMessage = "Cannot parse line."
	.ParseErrorTitle = .ScriptName
	.BankCode = ""
	.AccountNum = ""		' default if not specified in file
	.BranchCode = ""		' default if not specified in file
	.AccountType = "CREDITCARD"
	.QuickenBankID = ""		' copied to INTU.BID if present
	.CurrencyCode = "EUR"	' default if not specified in file
	.ColumnHeadersPresent = True	' are the column headers in the file?
	.SkipHeaderLines = 0	' number of lines to skip before the transaction data
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
	.CSVSeparator = ","
	.DecimalSeparator = ","	' as used in amounts
	.TxnLinePattern = ""
	.DateSequence = "DMY"	' must be DMY, MDY, or YMD
	.DateSeparator = "-/. "	' can be empty for dates in e.g. "yyyymmdd" format
	.OldestLast = False		' True if transactions are in reverse order
	.InvertSign = False	' make credits into debits etc
	.NoAvailableBalance = True		' True if file does not contain "Available Balance" information
	.MemoChunkLength = 0	' if memo field consists of fixed length chunks
	.TxnDatePattern = ".*(\d\d)\.(\d\d)\.(\d\d)\ (\d\d)\.(\d\d)"	' pattern to find transaction date in the memo
	.TxnDateSequence = Array(3,2,1,4,5,0)	' order of the info in the pattern (from 1 to 6): Y,M,D,H,M,S
	.PayeeLocation = 0		' start of payee in memo
	.PayeeLength = 0		' length of payee in memo
	.MonthNames = Array("jan","feb","mrt","apr","mei","jun","jul","aug","sep","okt","nov","dec", _
		"januari","februari","maart","april","mei","juni","juli","augustus","september","oktober","november","december")
'04 jun 05 jun HAVEN HOLIDAYS HEMEL HEMPSTE GB 752,00 GBP 969,85 debet
	.Fields = Array( _
		Array(fldBookDate, "BookDate"), _
		Array(fldTransactionDate, "TxnDate"), _
		Array(fldPayee, "Payee"), _
		Array(fldMemo, "Place"), _
		Array(fldMemo, "Country"), _
		Array(fldAmount, "Amount"), _
		Array(fldSkip, "DRCR", "debet|credit") _
	)
	
' min/max fields expected: default to size of Fields array. can be overridden here if required
'	.MinFieldsExpected = 1
'	.MaxFieldsExpected = 1
	.Properties = Array( _
		Array("AcctNum", "Account number", _
			"The account number for " & ScriptName, _
			ptString,,"=CheckAccount", "Please enter a valid account number.") _
		)
'	Set .TransactionCallback = GetRef("TransactionCallback")
'	Set .PreParseCallback = GetRef("PreParseCallback")
'	Set .HeaderCallback = GetRef("HeaderCallback")
'	Set .CustomDateCallback = GetRef("CustomDateCallback")
'	Set .CustomAmountCallback = GetRef("CustomAmountCallback")
'	Set .ReadLineCallback = GetRef("ReadLineCallback")
'	Set .FinaliseCallback = GetRef("FinaliseCallback")
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

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
	Dim sTmp
	sTmp = ReadLine()
	RecogniseTextFile = (sTmp = "International Card Services BV")
	If RecogniseTextFile Then
		LogProgress ScriptName, "File Recognised"
	End If
End Function

Dim sPatCC: sPatCC = "[A-Z][A-Z]"
Dim sPatCcy: sPatCcy = "[A-Z][A-Z][A-Z]"
Dim sPatAmt: sPatAmt = "[\d\.]+,\d\d"
Dim sPatDate: sPatDate = "\d\d [a-z][a-z][a-z]"
Function PatternReplace(sPat)
	Dim sTmp
	sTmp = Replace(sPat, "%date", sPatDate)
	sTmp = Replace(sTmp, "%amt", sPatAmt)
	sTmp = Replace(sTmp, "%ccy", sPatCcy)
	sTmp = Replace(sTmp, "%cc", sPatCC)
	PatternReplace = sTmp
End Function

Dim Stmt
Dim da: Set da = New DateAccumulator
Dim sDateYear

'30 mei 30 mei GEINCASSEERD VORIG 796,01 credit
'(\d\d mmm) (\d\d mmm) (merchant) (place) (CC) ((amt CCC)) (amt) (debet|credit)
Function LoadTextFile()
	Dim re1, re2: Set re1 = New RegExp: Set re2 = New RegExp
	Dim iTmp
	re1.Pattern = PatternReplace("(%date) (%date) (.*) (%cc(?: ))?(%amt) (debet|credit)")
	re2.Pattern = PatternReplace("(%date) (%date) (.*) (%cc) (%amt %ccy) (%amt) (debet|credit)")
	Dim iLine: iLine=0
	Dim sLine
	Dim sDC
	Dim vFields
	
	Set Stmt = NewStatement()
	Stmt.OpeningBalance.Ccy = "EUR"
	Stmt.ClosingBalance.Ccy = "EUR"
	Stmt.BankName = Params.BankCode
	Stmt.BranchName = Params.BranchCode
	Stmt.AcctType = Params.AccountType

	Do While Not AtEof()
		sLine = ReadLine()
		iLine = iLine + 1
		Select Case iLine
		Case 5
			' debet|credit debet|credit
			' for opening balance + payments received
			If StartsWith(sLine, "Datum") Then
' if no payments/transactions, 0,00 is neither debit nor credit and the text is omitted
				iLine = iLine + 1
				sDC = "credit credit"
			Else
				sDC = sLine
			End If
		Case 10
			' statement date = closing balance date
			Stmt.ClosingBalance.BalDate = Params.ParseDate(sLine)
'			MsgBox "clos bal " & sLine & " = " & Stmt.closingbalance.baldate
			da.Process(Stmt.ClosingBalance.BalDate)
			sDateYear = " " & CStr(Year(Stmt.ClosingBalance.BalDate))
		Case 11
			' opening balance
			Stmt.OpeningBalance.Amt = Params.ParseAmount(sLine)
			If Left(sDC, 1) = "d" Then Stmt.OpeningBalance.Amt = -Stmt.OpeningBalance.Amt
'			MsgBox "opening balance: " & Stmt.openingbalance.amt
		Case 12
			' account number
			Stmt.Acct = Trim(sLine)
		Case 15
			' new charges + closing balance
			iTmp = InStr(2, sLine, "€")
			Stmt.ClosingBalance.Amt = Params.ParseAmount(Mid(sLine,iTmp))
		Case 29
			' debet|credit debet|credit
			' for new charges + closing balance
			If sLine = "1" Then
				iLine = iLine + 1
			Else
				iTmp = InStr(sLine, " ")
				If Mid(sLine, iTmp+1, 1) = "d" Then Stmt.ClosingBalance.Amt = -Stmt.ClosingBalance.Amt
			End If
'			MsgBox "closing balance: " & Stmt.closingbalance.amt
		Case Else
			If iLine > 30 and IsNumeric(Left(sLine, 2)) Then
				If re2.Test(sLine) Then
'					MsgBox sLine & " matches international pattern"
					vFields = ParseLineFixed(sLine, re2.Pattern)
					ProcessIntl(vFields)
				Else
					If re1.Test(sLine) Then
'						MsgBox sLine & " matches domestic pattern"
						vFields = ParseLineFixed(sLine, re1.Pattern)
						ProcessDomestic(vFields)
'					Else
'						MsgBox sLine & " doesn't match"
					End If
				End If
			End if
		End Select
	Loop
	Stmt.OpeningBalance.BalDate = da.MinDate
	LoadTextFile = True
End Function

Function LoseLastWord(sPayee)
	Dim v: v=Split(sPayee, " ")
	Dim sTmp, i
	sTmp = v(LBound(v))
	For i=LBound(v)+1 To UBound(v)-1
		sTmp = sTmp & " " & v(i)
	Next
	LoseLastWord = sTmp
End Function

' (%date) (%date) (.*) (%cc) (%amt) (debet|credit)
Function ProcessDomestic(vFields)
	If IsEmpty(vFields) Then
		Exit Function
	End If
	NewTransaction
	Txn.TxnDate = MostRecent(Params.ParseDate(vFields(1) & sDateYear))
	Txn.TxnDateValid = (Txn.TxnDate <> NODATE)
	Txn.BookDate = MostRecent(Params.ParseDate(vFields(2) & sDateYear))
	da.Process Txn.BookDate
	Txn.Payee = LoseLastWord(vFields(3))
	Txn.Memo = Trim(vFields(3) & " " & vFields(4))
	ConcatMemo FormatDateTime(Txn.TxnDate, vbShortDate)
	Txn.Amt = Params.ParseAmount(vFields(5))
If txn.amt=0 Then MsgBox "zero amount: " & vfields(5)
	If vFields(6) = "debet" Then
		Txn.Amt = -Txn.Amt
	End If
	If Txn.Amt >= 0 Then
		Txn.TxnType = "DEP"
	Else
		Txn.TxnType = "PAYMENT"
	End If
End Function

' (%date) (%date) (.*) (%cc) (%amt %ccy) (%amt) (debet|credit)
Function ProcessIntl(vFields)
	Dim sTmp
	If IsEmpty(vFields) Then
		Exit Function
	End If
	NewTransaction
	sTmp = vFields(1)
'	MsgBox "Txn date (" & sTmp & ")=" & Params.ParseDate(sTmp & sDateYear)
	Txn.TxnDate = MostRecent(Params.ParseDate(vFields(1) & sDateYear))
	Txn.TxnDateValid = (Txn.TxnDate <> NODATE)
	Txn.BookDate = MostRecent(Params.ParseDate(vFields(2) & sDateYear))
	da.Process Txn.BookDate
	Txn.Payee = LoseLastWord(vFields(3))
	Txn.Memo = vFields(3) & " " & vFields(4)
	ConcatMemo FormatDateTime(Txn.TxnDate, vbShortDate)
	ConcatMemo vFields(5)
	Txn.Amt = Params.ParseAmount(vFields(6))
	If vFields(7) = "debet" Then
		Txn.Amt = -Txn.Amt
	End If
	If Txn.Amt >= 0 Then
		Txn.TxnType = "DEP"
	Else
		Txn.TxnType = "PAYMENT"
	End If
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
'MsgBox "In finalisation callback"
	Stmt.OpeningBalance.BalDate = da.MinDate
	FinaliseCallback = True
End Function
