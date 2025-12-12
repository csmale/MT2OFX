' MT2OFX Input Processing Script Argenta NL TSV (.tab) format

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/ArgentaNL-TSV.vbs 11    25/11/08 22:14 Colin $"

Dim Params: Set Params = New MT2OFXScript

With Params
	.MinimumProgramVersion = "3"
	.DebugRecognition = False	' enables debug code in recognition
	.ScriptName = "ArgentaNL2-TSV"
	.FormatName = "Argenta NL TSV"
	.ParseErrorMessage = "Cannot parse line."
	.ParseErrorTitle = .ScriptName
	.BankCode = "ARSNNL21"
	.AccountNum = ""		' default if not specified in file
	.BranchCode = ""		' default if not specified in file
	.QuickenBankID = ""		' copied to INTU.BID if present
	.CurrencyCode = "EUR"	' default if not specified in file
	.ColumnHeadersPresent = True	' are the column headers in the file?
	.SkipHeaderLines = 0	' number of lines to skip before the transaction data
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
	.CSVSeparator = vbTab
	.DecimalSeparator = ","	' as used in amounts
	.TxnLinePattern = ""
	.DateSequence = "DMY"	' must be DMY, MDY, or YMD
	.DateSeparator = "-/. "	' can be empty for dates in e.g. "yyyymmdd" format
	.OldestLast = True		' True if transactions are in reverse order
	.InvertSign = False	' make credits into debits etc
	.NoAvailableBalance = True		' True if file does not contain "Available Balance" information
	.MemoChunkLength = 0	' if memo field consists of fixed length chunks
	.TxnDatePattern = ".*TRANSACTIEDATUM\* (\d\d)-(\d\d)-(\d\d\d\d)"	' pattern to find transaction date in the memo
	.TxnDateSequence = Array(3,2,1,0,0,0)	' order of the info in the pattern (from 1 to 6): Y,M,D,H,M,S
	.PayeeLocation = 0		' start of payee in memo
	.PayeeLength = 0		' length of payee in memo
	.MonthNames = Empty
' "Rekeningnummer"	"Naam"	"Product type"	"Saldo"	"Valuta"	"Boekdatum"	"Valuta datum"
' "Rekeningnummer tegenpartij"	"Naam tegenpartij"	"Woonplaats tegenpartij"	"Bedrag"	"Valuta"
' "Omschrijving"	"Budgetcode"	"Naam begunstigde"	"Ident. code begunstigde"	"Ident. code opdrachtgever"
' "Betalingskenmerk"
	.Fields = Array( _
		Array(fldAccountNum, "Rekeningnummer", "=(""%1""=""Rekeningnummer"") Or Validate_Netherlands(""%1"")"), _
		Array(fldSkip, "Naam"), _
		Array(fldSkip, "Product type"), _
		Array(fldClosingBal, "Saldo"), _
		Array(fldCurrency, "Valuta", "Valuta|EUR"), _
		Array(fldBookDate, "Boekdatum"), _
		Array(fldValueDate, "Valuta datum"), _
		Array(fldSkip, "Rekeningnummer tegenpartij"), _
		Array(fldPayee, "Naam tegenpartij"), _
		Array(fldSkip, "Woonplaats tegenpartij"), _
		Array(fldAmount, "Bedrag"), _
		Array(fldSkip, "Valuta"), _
		Array(fldMemo, "Omschrijving"), _
		Array(fldSkip, "Budgetcode"), _
		Array(fldSkip, "Naam begunstigde"), _
		Array(fldSkip, "Ident. code begunstigde"), _
		Array(fldSkip, "Ident. code opdrachtgever"), _
		Array(fldSkip, "Betalingskenmerk") _
	)
' min/max fields expected: default to size of Fields array. can be overridden here if required
' Argenta defines 18 columns but only uses the first 15 of them
	.MinFieldsExpected = 15
	.MaxFieldsExpected = 19
	.Properties = Array()
	Set .TransactionCallback = GetRef("TransactionCallback")
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
	Dim sLine
' older files did not have the header line
	sLine = ReadLine()
	If InStr(sLine, vbTab & """Product type""" & vbTab) = 0 Then
		Params.ColumnHeadersPresent = False
	End If
	Rewind()
	RecogniseTextFile = DefaultRecogniseTextFile(Params)
	If RecogniseTextFile Then
		LogProgress ScriptName, "File Recognised"
	End If
End Function

Function LoadTextFile()
	LoadTextFile = DefaultLoadTextFile(Params)
End Function

' callback functions, called from DefaultRecogniseTextFile and DefaultLoadTextFile
' ths following implementations are functionally neutral or equivalent to the default processing
' in the class
Function HeaderCallback(sLine)
MsgBox "In header callback: " & sLine
	HeaderCallback = True
End Function

Function TransactionCallback(t, vFields)
' MsgBox "In transaction callback: " & t.Memo
	Dim sTmp, iTmp
	
' memo and payee sometimes have XML entities...
	t.Memo = Replace(t.Memo, "&gt;", ">")
	t.Memo = Replace(t.Memo, "&lt;", "<")
	t.Memo = Replace(t.Memo, "&amp;", "&")

	t.Payee = Replace(t.Payee, "&gt;", ">")
	t.Payee = Replace(t.Payee, "&lt;", "<")
	t.Payee = Replace(t.Payee, "&amp;", "&")
	
' cash withdrawals - special string for payee
	If StartsWith(t.Memo, "GELDAUTOMAAT ") Then
		t.Payee = "Kasopname"
	Elseif StartsWith(t.Memo, "CHIPKNIP ") Then
		t.Payee = "Chipknip"
	Elseif StartsWith(t.Memo, "RENTE") Then
		t.Payee = "Rente"
	End If
			
' sometimes we don't get a payee!
	If t.Payee = "" Then
		If StartsWith(t.Memo, "BETAALAUTOMAAT ") Then
			sTmp = Mid(t.Memo, 16)	' lose BETAALAUTOMAAT
			sTmp = Left(sTmp, Len(sTmp)-19)	' lose date/time
			sTmp = Trim(Left(sTmp, 32))		' OFX: max len 32
			t.Payee = sTmp
		Else
			t.Payee = "Onbekend"
		End If
	End If

' transaction type
	If StartsWith(t.Memo, "GELDAUTOMAAT ") _
	Or StartsWith(t.Memo, "CHIPKNIP OP") Then
		t.TxnType = "ATM"
	Elseif StartsWith(t.Memo, "BETAALAUTOMAAT ") Then
		t.TxnType = "POS"
	Elseif StartsWith(t.Memo, "RENTE") Then
		t.TxnType = "INT"
	End If

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
