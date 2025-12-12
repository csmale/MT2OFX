' MT2OFX Input Processing Script Citibank (US) CSV format

Option Explicit

Const ScriptVersion = "$Header$"

Dim Params: Set Params = New MT2OFXScript

With Params
	.MinimumProgramVersion = "3"
	.DebugRecognition = False	' enables debug code in recognition
	.ScriptName = "CitibankUS-CSV"
	.FormatName = "Citibank (US) CSV"
	.ParseErrorMessage = "Cannot parse line."
	.ParseErrorTitle = .ScriptName
	.BankCode = "CITIUS33"
	.AccountNum = "in file"		' default if not specified in file
	.BranchCode = ""		' default if not specified in file
	.AccountType = "CREDITCARD"
	.QuickenBankID = ""		' copied to INTU.BID if present
	.CurrencyCode = "USD"	' default if not specified in file
	.ColumnHeadersPresent = True	' are the column headers in the file?
	.SkipHeaderLines = 12	' number of lines to skip before the transaction data
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
	.CSVSeparator = ","
	.DecimalSeparator = "."	' as used in amounts
	.TxnLinePattern = ""
	.DateSequence = "MDY"	' must be DMY, MDY, or YMD
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
' Date,Type,BAI Code,Bank Ref#,Customer Ref#,Amount
	.Fields = Array( _
		Array(fldBookDate, "Date"), _
		Array(fldSkip, "Type"), _
		Array(fldSkip, "BAI Code"), _
		Array(fldMemo, "Bank Ref#"), _
		Array(fldMemo, "Customer Ref#"), _
		Array(fldAmount, "Amount") _
	)
' min/max fields expected: default to size of Fields array. can be overridden here if required
'	.MinFieldsExpected = 1
'	.MaxFieldsExpected = 1
'	.Properties = Array( _
'		Array("AcctNum", "Account number", _
'			"The account number for " & Params.FormatName, _
'			ptString,,"=CheckAccount", "Please enter a valid account number.") _
'		)
	Set .TransactionCallback = GetRef("TransactionCallback")
	Set .IsValidTxnLine = GetRef("IsValidTxnLine")
'	Set .PreParseCallback = GetRef("PreParseCallback")
	Set .HeaderCallback = GetRef("HeaderCallback")
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

Dim tLast: Set tLast = Nothing

Function HeaderCallback(sLine)
'MsgBox "In header callback: " & sLine
	Dim sTmp, iTmp
	iTmp = InStr(sLine, Params.CSVSeparator)
	If iTmp=0 Then
		sTmp = sLine
	Else
		sTmp = Trim(Left(sLine, iTmp-1))
	End If
	If StartsWith(sTmp, "Bank ABA#: ") Then
		Params.BranchCode = Mid(sTmp, 12)
	ElseIf StartsWith(sTmp, "Account#: ") Then
		sTmp = Mid(sTmp, 11)
		iTmp = InStr(sTmp, " ")
		If iTmp>0 Then sTmp = Left(sTmp, iTmp-1)
		Params.AccountNum = sTmp
	End If
	HeaderCallback = True
End Function

Function TransactionCallback(t, vFields)
'MsgBox "In transaction callback: " & t.Memo
	Dim sTmp, iTmp, sLine, sMemo, sOldMemo
	sOldMemo = t.Memo
	Do Until InStr(sLine, ",") > 0
		sLine = ReadLine()
		iTmp = InStr(sLine, ",")
		If iTmp>0 Then
			sTmp = Trim(Left(sLine, iTmp-1))
		Else
			sTmp = Trim(sLine)
		End If
		If Len(sMemo) = 0 Then
			sMemo = sTmp
		Else
			sMemo = sMemo & Cfg.MemoDelimiter & sTmp
		End If
	Loop
	If Left(sMemo, 1) = """" Then
		sMemo = Trim(Mid(sMemo, 2, Len(sMemo)-2))
	End If
	t.Memo = sMemo
	ConcatMemo sOldMemo
	sTmp = Bia2Ofx(vFields(3))
	If Len(sTmp) > 0 Then
		t.TxnType = sTmp
	End If
	TransactionCallback = True
End Function
Function CustomDateCallback(sDate)
MsgBox "In custom date callback: " & sDate
	CustomDateCallback = ParseDateEx(sDate, Params.DateSequence, Params.DateSeparator)
End Function
Function CustomAmountCallback(sAmt)
	Dim sTmp
	If InStr(sAmt, "(") > 0 Then
		sTmp = "-" & Replace(Replace(sAmt, "(", ""), ")", "")
	Else
		sTmp = sAmt
	End If
	CustomAmountCallback = ParseNumber(sTmp, Params.DecimalSeparator)
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
'MsgBox "In IsValidTxnLine callback: " & sLine
	IsValidTxnLine = (IsNumeric(Left(sLine, 1)))
End Function

Function Bia2Ofx(sBia)
	Dim iBia, sTmp
	If Not IsNumeric(sBia) Then
		Exit Function
	End If
	iBia = CLng(sBia)
	Select Case iBia
	Case 195, 275, 508, 577: sTmp = "XFER"
	Case 175, 475: sTmp = "CHECK"
	Case 698: sTmp = "SRVCHG"
	End Select
	Bia2Ofx = sTmp
End Function
