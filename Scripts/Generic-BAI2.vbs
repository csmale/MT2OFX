' MT2OFX Input Processing Script Generic BAI2 format

Option Explicit

Const ScriptVersion = "$Header$"

Dim Params: Set Params = New MT2OFXScript

With Params
	.MinimumProgramVersion = "3"
	.DebugRecognition = False	' enables debug code in recognition
	.ScriptName = "Generic-BAI2"
	.FormatName = "Generic BAI2 Format"
	.ParseErrorMessage = "Cannot parse line."
	.ParseErrorTitle = .ScriptName
	.BankCode = ""
	.AccountNum = ""		' default if not specified in file
	.BranchCode = ""		' default if not specified in file
	.QuickenBankID = ""		' copied to INTU.BID if present
	.CurrencyCode = "USD"	' default if not specified in file
	.ColumnHeadersPresent = False	' are the column headers in the file?
	.SkipHeaderLines = 0	' number of lines to skip before the transaction data
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
	.CSVSeparator = ","
	.DecimalSeparator = "."	' as used in amounts
	.TxnLinePattern = ""
	.DateSequence = "YMD"	' must be DMY, MDY, or YMD
	.DateSeparator = ""		' can be empty for dates in e.g. "yyyymmdd" format
	.OldestLast = True		' True if transactions are in reverse order
	.InvertSign = False	' make credits into debits etc
	.NoAvailableBalance = True		' True if file does not contain "Available Balance" information
	.MemoChunkLength = 0	' if memo field consists of fixed length chunks
	.TxnDatePattern = ".*(\d\d)\.(\d\d)\.(\d\d)\ (\d\d)\.(\d\d)"	' pattern to find transaction date in the memo
	.TxnDateSequence = Array(3,2,1,4,5,0)	' order of the info in the pattern (from 1 to 6): Y,M,D,H,M,S
	.PayeeLocation = 0		' start of payee in memo
	.PayeeLength = 0		' length of payee in memo
	.MonthNames = Empty
' 16,142,5346,,,,WELLCAREFLCAID    PAYMENT           090416/
	.Fields = Array( _
		Array(fldSkip, "Record Code", "16"), _
		Array(fldSkip, "Type Number", "\d+"), _
		Array(fldAmount, "Amount"), _
		Array(fldSkip, "Funds Type", "[012SVDZ]?"), _
		Array(fldSkip, "Bank reference"), _
		Array(fldSkip, "Customer reference"), _
		Array(fldMemo, "Text"), _
		Array(fldSkip, ""), _
		Array(fldSkip, ""), _
		Array(fldSkip, ""), _
		Array(fldSkip, "") _
	)
' min/max fields expected: default to size of Fields array. can be overridden here if required
	.MinFieldsExpected = 7
	.MaxFieldsExpected = 11
	.Properties = Array( _
		Array("AcctNum", "Account number", _
			"The account number for " & Params.FormatName, _
			ptString,,"=CheckAccount", "Please enter a valid account number.") _
		)
	Set .MyReadLine = GetRef("BAI2ReadLine")
	Set .MyAtEOF = GetRef("BAI2AtEof")
	Set .MyRewind = GetRef("BAI2Rewind")
'	Set .TransactionCallback = GetRef("TransactionCallback")
	Set .IsValidTxnLine = GetRef("IsValidTxnLine")
	Set .PreParseCallback = GetRef("PreParseCallback")
'	Set .HeaderCallback = GetRef("HeaderCallback")
'	Set .StatementCallback = GetRef("StatementCallback")
'	Set .CustomDateCallback = GetRef("CustomDateCallback")
	Set .CustomAmountCallback = GetRef("CustomAmountCallback")
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
	RecogniseTextFile = False
	sLine = ReadLine()
	If Left(sLine, 3) <> "01," Then Exit Function
	If Right(sLine, 1) <> "/" Then Exit Function
' 01,071000152,493172,090420,1037,1,80,,2/
	Dim vFields
	vFields = ParseLineDelimited(sLine, ",")
'msgbox "Ubound: " & UBound(vFields) & "; date=" & vFields(4)
	If UBound(vFields) <> 9 Then Exit Function
	If Not StringMatches(vFields(2), ".+") Then Exit Function
	If Not StringMatches(vFields(3), ".+") Then Exit Function
	If Not StringMatches(vFields(4), "\d{6}") Then Exit Function
	If Not StringMatches(vFields(5), "\d{4}") Then Exit Function
	If Not StringMatches(vFields(6), "\d*") Then Exit Function
	If Not StringMatches(vFields(7), "\d*") Then Exit Function
	If Not StringMatches(vFields(8), "\d*") Then Exit Function
	If Not StringMatches(vFields(9), "2") Then Exit Function
	
	RecogniseTextFile = True
	
	If RecogniseTextFile Then
		LogProgress ScriptName, "File Recognised"
	End If
End Function

Function LoadTextFile()
	Dim sAcct
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
MsgBox "In transaction callback: " & t.Memo
	TransactionCallback = True
End Function
Function CustomDateCallback(sDate)
MsgBox "In custom date callback: " & sDate
	CustomDateCallback = ParseDateEx(sDate, Params.DateSequence, Params.DateSeparator)
End Function
Function CustomAmountCallback(sAmt)
'MsgBox "In custom amount callback: " & sAmt
	CustomAmountCallback = ParseNumber(sAmt, Params.DecimalSeparator) / 100.0
End Function
Function ReadLineCallback(sLine)
MsgBox "In read line callback: " & sLine
	ReadLineCallback = sLine
End Function
' PreParseCallback: returns True or False. True means the line can be processed; False means skip this line.
Function PreParseCallback(vFields)
'MsgBox "In preparse callback: " & UBound(vFields) & " fields."
	Dim i
	If vFields(1) = "16" Then
		For i=8 To UBound(vFields)
			vFields(7) = vFields(7) & "," & vFields(i)
			vFields(i) = ""
		Next
	End If
	PreParseCallback = True
End Function
Function FinaliseCallback()
MsgBox "In finalisation callback"
	FinaliseCallback = True
End Function
Function IsValidTxnLine(sLine)
'MsgBox "In IsValidTxnLine callback: " & sLine
	IsValidTxnLine = Left(sLine, 3) = "16,"
End Function

' sLineBuf holds the next physical line
Dim sLineBuf: sLineBuf = ""
Function BAI2ReadLine()
	Dim sLine
	If Len(sLineBuf) > 0 Then
		sLine = sLineBuf
		sLineBuf = ""
	Else
		sLine = ReadLine()
	End If
	If Right(sLine, 1) = "/" Then sLine = Left(sLine, Len(sLine)-1)
	Do While Not AtEof()
		sLineBuf = ReadLine()
		If Right(sLineBuf, 1) = "/" Then sLineBuf = Left(sLineBuf, Len(sLineBuf)-1)
		If Left(sLineBuf, 3) = "88," Then
			sLine = sLine & Mid(sLineBuf, 4)
		Else
			Exit Do
		End If
	Loop
' danger: type 16 records (actual transactions) can contain commas in the text itself! these must be escaped
' 16,142,11612972,,,,FCSO, INC.        MED B PAY                           705TRN/
' 88,*1*884684308/
	If Left(sLine, 3) = "16," Then
		
	End If
	BAI2ReadLine = sLine
End Function
Function BAI2AtEof()
	BAI2AtEof = (Len(sLineBuf) = 0) And AtEof()
End Function
Function BAI2Rewind()
	Rewind
	sLineBuf = ""
End Function
