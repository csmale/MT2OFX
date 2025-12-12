' MT2OFX Input Processing Script Fortis Bank (NL) CSV format

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/FortisNL-CSV.vbs 1     10/02/10 0:03 Colin $"

Dim Params: Set Params = New MT2OFXScript

With Params
	.MinimumProgramVersion = "3"
	.DebugRecognition = False	' enables debug code in recognition
	.ScriptName = "FortisNL-CSV"
	.FormatName = "Fortis Bank (NL) CSV"
	.ParseErrorMessage = "Cannot parse line."
	.ParseErrorTitle = .ScriptName
	.BankCode = "FTSBNL22"
	.AccountNum = ""		' default if not specified in file
	.BranchCode = ""		' default if not specified in file
	.AccountType = "CHECKING"	' can be CHECKING or CREDITCARD
	.QuickenBankID = ""		' copied to INTU.BID if present
	.CurrencyCode = "EUR"	' default if not specified in file
	.ColumnHeadersPresent = True	' are the column headers in the file?
	.SkipHeaderLines = 0	' number of lines to skip before the transaction data
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
	.CSVSeparator = ";"
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
' REKENINGNUMMER;TEN LASTE VAN;BOEKDATUM;VALUTADATUM;BEDRAG;MUNT V/D REKENING;BEGUNSTIGDE;BEGUNSTIGDE REKENING;MUTATIE OMSCHRIJVING;REGEL 1;REGEL 2;REGEL 3;REGEL 4;REGEL 5
' 12.20.66.123;PIETERSEN;13-1-2009;13-1-2009;-6,85;EUR;;;BETAALAUTOMAAT;BETAALD  13-01-09 11U09 74EV01 ;ZEEMAN textielSupers>AMSTELVEEN> ;PASNR 071 ; ; 

	.Fields = Array( _
		Array(fldAccountNum, "REKENINGNUMMER"), _
		Array(fldSkip, "TEN LASTE VAN"), _
		Array(fldBookDate, "BOEKDATUM"), _
		Array(fldValueDate, "VALUTADATUM"), _
		Array(fldAmount, "BEDRAG"), _
		Array(fldCurrency, "MUNT V/D REKENING"), _
		Array(fldPayee, "BEGUNSTIGDE"), _
		Array(fldSkip, "BEGUNSTIGDE REKENING"), _
		Array(fldMemo, "MUTATIE OMSCHRIJVING"), _
		Array(fldMemo, "REGEL 1"), _
		Array(fldMemo, "REGEL 2"), _
		Array(fldMemo, "REGEL 3"), _
		Array(fldMemo, "REGEL 4"), _
		Array(fldMemo, "REGEL 5") _	
	)
' min/max fields expected: default to size of Fields array. can be overridden here if required
'	.MinFieldsExpected = 1
'	.MaxFieldsExpected = 1
	.Properties = Array( _
		Array("AcctNum", "Account number", _
			"The account number for " & Params.FormatName, _
			ptString,,"=CheckAccount", "Please enter a valid account number.") _
		)
	Set .TransactionCallback = GetRef("TransactionCallback")
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
	RecogniseTextFile = DefaultRecogniseTextFile(Params)
	If RecogniseTextFile Then
		LogProgress ScriptName, "File Recognised"
	End If
End Function

Function LoadTextFile()
	Dim sAcct
	If Not Params.FieldDict.Exists(fldAccountNum) Then
		If PropertyExists("AcctNum") Then
			sAcct = GetProperty("AcctNum")
			If Len(sAcct) = 0 Then
				Message True, True, "This file does not contain an account number. Please set the Account number through Options, Scripts, Parameters.", Params.ScriptTitle
				LoadTextFile = False
				Exit Function
			End If
			Params.AccountNum = sAcct
		Else
			If Len(Params.AccountNum) = 0 Then
				Message True, True, "Script error: no account number as constant, field or property", Params.ScriptTitle
				LoadTextFile = False
				Exit Function
			End If
		End If
	End If
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
'MsgBox "In transaction callback: " & t.Memo
    Dim iTmp
    If t.Payee = "" Then
        t.Payee = vFields(11)
        iTmp = InStr(t.Payee, ">")
        If iTmp > 0 Then
            t.Payee = Trim(Left(t.Payee, iTmp-1))
        End If
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
