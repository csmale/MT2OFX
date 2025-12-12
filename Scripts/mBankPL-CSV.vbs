' MT2OFX Input Processing Script mBank (Poland) CSV format
' NB: Must be saved in UTF-8

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/mBankPL-CSV.vbs 1     1/06/09 21:58 Colin $"

Dim Params: Set Params = New MT2OFXScript

With Params
	.MinimumProgramVersion = "3"
	.DebugRecognition = False	' enables debug code in recognition
	.ScriptName = "mBankPL-CSV"
	.FormatName = "mBank (Poland) CSV"
	.ParseErrorMessage = "Cannot parse line."
	.ParseErrorTitle = .ScriptName
	.BankCode = "BREXPLPWMBK"
	.AccountNum = "in header"		' default if not specified in file
	.BranchCode = ""		' default if not specified in file
	.QuickenBankID = ""		' copied to INTU.BID if present
	.CurrencyCode = ""	' default if not specified in file
	.ColumnHeadersPresent = True	' are the column headers in the file?
	.SkipHeaderLines = 44	' number of lines to skip before the transaction data
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
	.CSVSeparator = ";"
	.DecimalSeparator = ","	' as used in amounts
	.TxnLinePattern = ""
	.DateSequence = "YMD"	' must be DMY, MDY, or YMD
	.DateSeparator = "-/. "	' can be empty for dates in e.g. "yyyymmdd" format
	.OldestLast = False		' True if transactions are in reverse order
	.InvertSign = False	' make credits into debits etc
	.NoAvailableBalance = True		' True if file does not contain "Available Balance" information
	.MemoChunkLength = 0	' if memo field consists of fixed length chunks
	.TxnDatePattern = ".*(\d\d)\.(\d\d)\.(\d\d)\ (\d\d)\.(\d\d)"	' pattern to find transaction date in the memo
	.TxnDateSequence = Array(3,2,1,4,5,0)	' order of the info in the pattern (from 1 to 6): Y,M,D,H,M,S
	.PayeeLocation = 0		' start of payee in memo
	.PayeeLength = 0		' length of payee in memo
	.MonthNames = Empty
' #Data operacji;#Data księgowania;#Opis operacji;#Nazwa;#Rachunek;#Nazwa Banku;#Opis dodatkowy;#Kwota;#Saldo po operacji
	.Fields = Array( _
		Array(fldBookDate, "#Data operacji"), _
		Array(fldValueDate, "#Data księgowania"), _
		Array(fldMemo, "#Opis operacji"), _
		Array(fldPayee, "#Nazwa"), _
		Array(fldSkip, "#Rachunek"), _
		Array(fldSkip, "#Nazwa Banku"), _
		Array(fldMemo, "#Opis dodatkowy"), _
		Array(fldAmount, "#Kwota"), _
		Array(fldClosingBal, "#Saldo po operacji"), _
		Array(fldSkip, "") _
	)
' min/max fields expected: default to size of Fields array. can be overridden here if required
	.MinFieldsExpected = 9
	.MaxFieldsExpected = 10
	.Properties = Array( _
		Array("PLNPLZ", "Use PLZ currency", _
			"Replace PLN with PLZ in output", _
			ptBoolean) _
		)
'	Set .TransactionCallback = GetRef("TransactionCallback")
'	Set .IsValidTxnLine = GetRef("IsValidTxnLine")
'	Set .PreParseCallback = GetRef("PreParseCallback")
	Set .HeaderCallback = GetRef("HeaderCallback")
	Set .StatementCallback = GetRef("StatementCallback")
'	Set .CustomDateCallback = GetRef("CustomDateCallback")
'	Set .CustomAmountCallback = GetRef("CustomAmountCallback")
'	Set .ReadLineCallback = GetRef("ReadLineCallback")
End With

Dim bPLNPLZ

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
	bPLNPLZ = GetProperty("PLNPLZ")
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
	Session.InputFile.CodePage = 1250
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
	Session.InputFile.CodePage = 1250
	LoadTextFile = DefaultLoadTextFile(Params)
End Function

' callback functions, called from DefaultRecogniseTextFile and DefaultLoadTextFile
' ths following implementations are functionally neutral or equivalent to the default processing
' in the class
Dim sLastHeader
Dim dClosingBal, dOpeningBal
Function HeaderCallback(sLine)
	Dim sTmp
'MsgBox "In header callback: " & sLine
	If Left(sLine, 1) = "#" Then
		sLastHeader = sLine
'		msgbox "new header: " & sline
	Else
'		msgbox "header data for " & slastheader & ": " & sline
		Select Case sLastHeader
		Case "#Waluta;"
			Params.CurrencyCode = Left(sLine, 3)
			If bPLNPLZ And Params.CurrencyCode = "PLN" Then
				Params.CurrencyCode = "PLZ"
			End If
		Case "#Nr rachunku;"
			sTmp = Left(Replace(sLine, " ", ""), 26)
			Params.AccountNum = Right(sTmp, 16)
			Params.BranchCode = Mid(sTmp, 3, 8)
		Case "#Saldo początkowe;"
			dClosingBal = Params.ParseAmount(sLine)
		Case "#Saldo końcowe;"
			dOpeningBal = Params.ParseAmount(sLine)
		End Select
	End If
	HeaderCallback = True
End Function
Function TransactionCallback(t, vFields)
MsgBox "In transaction callback: " & t.Memo
	TransactionCallback = True
End Function
Function StatementCallback(s)
	s.OpeningBalance.Amt = dOpeningBal
	s.ClosingBalance.Amt = dClosingBal
	StatementCallback = True
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
Function IsValidTxnLine(sLine)
MsgBox "In IsValidTxnLine callback: " & sLine
	IsValidTxnLine = (IsNumeric(Left(sLine, 1)))
End Function
