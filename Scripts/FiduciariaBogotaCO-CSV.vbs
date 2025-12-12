' MT2OFX Input Processing Script Bank (Colombia) CSV format

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/FiduciariaBogotaCO-CSV.vbs 3     13/09/08 8:48 Colin $"

Dim Params: Set Params = New MT2OFXScript

With Params
	.MinimumProgramVersion = "3.5"
	.DebugRecognition = False	' enables debug code in recognition
	.ScriptName = "FiduciariaBogotaCO-CSV"
	.FormatName = "Fiduciaria Bogota (Colombia) CSV"
	.ParseErrorMessage = "Cannot parse line."
	.ParseErrorTitle = .ScriptName
	.BankCode = "FBOGCOB1"
	.AccountNum = ""		' default if not specified in file
	.BranchCode = ""		' default if not specified in file
	.QuickenBankID = ""		' copied to INTU.BID if present
	.CurrencyCode = "COP"	' default if not specified in file
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
	.NoAvailableBalance = True		' True if file does not contain "Available Balance" information
	.MemoChunkLength = 0	' if memo field consists of fixed length chunks
	.TxnDatePattern = ".*(\d\d)\.(\d\d)\.(\d\d)\ (\d\d)\.(\d\d)"	' pattern to find transaction date in the memo
	.TxnDateSequence = Array(3,2,1,4,5,0)	' order of the info in the pattern (from 1 to 6): Y,M,D,H,M,S
	.PayeeLocation = 0		' start of payee in memo
	.PayeeLength = 0		' length of payee in memo
	.MonthNames = Empty
' Fecha,Transacción,Oficina,Documento,Débito,Crédito, Efectivo, Cheque,
	.Fields = Array( _
		Array(fldBookDate, "Fecha"), _
		Array(fldMemo, "Transacción"), _
		Array(fldMemo, "Oficina"), _
		Array(fldPayee, "Documento"), _
		Array(fldAmtDebit, "Débito"), _
		Array(fldAmtCredit, "Crédito"), _
		Array(fldSkip, "Efectivo"), _
		Array(fldSkip, "Cheque"), _
		Array(fldSkip, "") _
	)
' min/max fields expected: default to size of Fields array. can be overridden here if required
	.MinFieldsExpected = 8
'	.MaxFieldsExpected = 1
	.Properties = Array( _
		Array("AcctNum", "Account number", _
			"The account number for " & .FormatName, _
			ptString,,"", "Please enter a valid account number.") _
		)
	Set .TransactionCallback = GetRef("TransactionCallback")
	Set .HeaderCallback = GetRef("HeaderCallback")
	Set .CustomDateCallback = GetRef("CustomDateCallback")
	Set .CustomAmountCallback = GetRef("CustomAmountCallback")
	Set .ReadLineCallback = GetRef("ReadLineCallback")
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
MsgBox "CheckAccount(""" & s & """) = " & CheckAccount
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
	Params.AccountNum = GetProperty("AcctNum")
	If Params.AccountNum = "" Then
		Call Configure
		Params.AccountNum = GetProperty("AcctNum")
		If Params.AccountNum = "" Then
			MsgBox "Sorry, must have an account number!"
			Abort
		End If
	End If
	LoadTextFile = DefaultLoadTextFile(Params)
End Function

' callback functions, called from DefaultRecogniseTextFile and DefaultLoadTextFile
' ths following implementations are functionally neutral or equivalent to the default processing
' in the class
Function HeaderCallback(sLine)
'MsgBox "In header callback: " & sLine
	HeaderCallback = True
End Function
Function TransactionCallback(t, vFields)
'MsgBox "In transaction callback: " & t.Memo
	TransactionCallback = True
End Function
Function CustomDateCallback(sDate)
'MsgBox "In custom date callback: " & sDate
	Dim dTmp
	Dim sTmp: sTmp = "2008/" & Trim(sDate)
	dTmp = ParseDateEx(sTmp, Params.DateSequence, Params.DateSeparator)
	If dTmp <> NODATE Then
		dTmp = MostRecent(dTmp)
	End If
	CustomDateCallback = dTmp
End Function
Function CustomAmountCallback(sAmt)
'MsgBox "In custom amount callback: " & sAmt
	CustomAmountCallback = ParseNumber(sAmt, Params.DecimalSeparator)
End Function
Function ReadLineCallback(sLine)
'MsgBox "In read line callback: " & sLine
	ReadLineCallback = sLine
End Function
Function FinaliseCallback()
'MsgBox "In finalisation callback"
	FinaliseCallback = True
End Function
