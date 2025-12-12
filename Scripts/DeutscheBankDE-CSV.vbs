' MT2OFX Input Processing Script Deutsche Bank credit card CSV format
' NB: This Script Will Not Work Without Customisation!

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/DeutscheBankDE-CSV.vbs 2     18/03/10 23:05 Colin $"

Dim Params: Set Params = New MT2OFXScript

With Params
	.MinimumProgramVersion = "3"
	.DebugRecognition = False	' enables debug code in recognition
	.ScriptName = "DeutscheBankCC-CSV"
	.FormatName = "Deutsche Bank Creditcard CSV"
	.ParseErrorMessage = "Cannot parse line."
	.ParseErrorTitle = .ScriptName
	.BankCode = "DEUTDEBB"
	.AccountNum = "InFile"		' default if not specified in file
	.BranchCode = ""		' default if not specified in file
	.AccountType = "CREDITCARD"	' can be CHECKING or CREDITCARD
	.QuickenBankID = ""		' copied to INTU.BID if present
	.CurrencyCode = "EUR"	' default if not specified in file
	.ColumnHeadersPresent = True	' are the column headers in the file?
	.SkipHeaderLines = 4	' number of lines to skip before the transaction data
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
	.CSVSeparator = ";"
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
' Voucher date;Date of receipt;Reason for payment;Foreign currency;Amount;Exchange rate;Amount;Currency
	.Fields = Array( _
		Array(fldTransactionDate, "Voucher date"), _
		Array(fldBookDate, "Date of receipt"), _
		Array(fldMemo, "Reason for payment"), _
		Array(fldMemo, "Foreign currency"), _
		Array(fldMemo, "Amount"), _
		Array(fldMemo, "Exchange rate"), _
		Array(fldAmount, "Amount"), _
		Array(fldCurrency, "Currency") _
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
	Set .StatementCallback = GetRef("StatementCallback")
'	Set .CustomDateCallback = GetRef("CustomDateCallback")
'	Set .CustomAmountCallback = GetRef("CustomAmountCallback")
	Set .ReadLineCallback = GetRef("ReadLineCallback")
End With

'Belegdatum;Eingangstag;Verwendungszweck;Fremdwährung;Betrag;Kurs;Betrag;Währung
Dim aFieldsGerman
aFieldsGerman = Array( _
		Array(fldTransactionDate, "Belegdatum"), _
		Array(fldBookDate, "Eingangstag"), _
		Array(fldMemo, "Verwendungszweck"), _
		Array(fldMemo, "Fremdwährung"), _
		Array(fldMemo, "Betrag"), _
		Array(fldMemo, "Kurs"), _
		Array(fldAmount, "Betrag"), _
		Array(fldCurrency, "Währung") _
)

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

' to hold the balance in the trailer line
Dim dBalance

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
    Dim sLine: sLine = ReadLine()
    If StartsWith(sLine, "Kreditkartentransaktionen") Then
        Params.Fields = aFieldsGerman
        Params.DateSequence = "DMY"
        Params.DecimalSeparator = ","
    End If
    Rewind
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
'MsgBox "In header callback: " & sLine
    Dim sTmp
    If StartsWith(sLine, "MasterCard") Then
        sTmp = Left(Replace(Mid(sLine, 12), " ", ""),16)
        Params.AccountNum = sTmp
    End If
    HeaderCallback = True
End Function
Function TransactionCallback(t, vFields)
'MsgBox "In transaction callback: " & t.Memo
    t.Payee = vFields(3)
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
'MsgBox "In read line callback: " & sLine
    If Left(sLine, 1) = """" And Right(sLine, 1) = """" Then
        sLine = Mid(sLine, 2, Len(sLine)-2)
    End If
    ReadLineCallback = sLine
End Function
' PreParseCallback: returns True or False. True means the line can be processed; False means skip this line.
Function PreParseCallback(vFields)
MsgBox "In preparse callback: " & UBound(vFields) & " fields."
	PreParseCallback = True
End Function
Sub StatementCallback(Stmt)
'MsgBox "In statement callback"
    Stmt.ClosingBalance.Amt = dBalance
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
    Dim vFields
' MsgBox "In IsValidTxnLine callback: " & sLine
    If IsNumeric(Left(sLine, 1)) Then
        IsValidTxnLine = txnlineNORMAL
    Else
        If Left(sLine, 6) = "Total:" Or Left(sLine, 6) = "Summe:" Then
            vFields = ParseLineDelimited(sLine, Params.CSVSeparator)
            If UBound(vFields) = 8 Then
                dBalance = ParseNumber(vFields(7), Params.DecimalSeparator)
            End If
        End If
        IsValidTxnLine = txnlineSKIP
    End If
End Function
