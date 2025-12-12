' MT2OFX Input Processing Script Basic CSV format
' NB: This Script Will Not Work Without Customisation!

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/CitibankMastercard-CSV.vbs 5     24/11/09 22:04 Colin $"

Dim Params: Set Params = New MT2OFXScript

With Params
	.MinimumProgramVersion = "3"
	.DebugRecognition = False	' enables debug code in recognition
	.ScriptName = "CitibankMastercard-CSV"
	.FormatName = "Citibank Mastercard CSV"
	.ParseErrorMessage = "Cannot parse line."
	.ParseErrorTitle = .ScriptName
	.BankCode = "CITIUS33"
	.AccountNum = "in file"		' default if not specified in file
	.BranchCode = ""		' default if not specified in file
	.AccountType = "CREDITCARD"
	.QuickenBankID = ""		' copied to INTU.BID if present
	.CurrencyCode = "USD"	' default if not specified in file
	.ColumnHeadersPresent = True	' are the column headers in the file?
	.SkipHeaderLines = 34	' number of lines to skip before the transaction data
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
	.CSVSeparator = ","
	.DecimalSeparator = "."	' as used in amounts
	.TxnLinePattern = ""
	.DateSequence = "DMY"	' must be DMY, MDY, or YMD
	.DateSeparator = "-"	' can be empty for dates in e.g. "yyyymmdd" format
	.OldestLast = True		' True if transactions are in reverse order
	.InvertSign = True	' make credits into debits etc
	.NoAvailableBalance = True		' True if file does not contain "Available Balance" information
	.MemoChunkLength = 0	' if memo field consists of fixed length chunks
	.TxnDatePattern = ".*(\d\d)\.(\d\d)\.(\d\d)\ (\d\d)\.(\d\d)"	' pattern to find transaction date in the memo
	.TxnDateSequence = Array(3,2,1,4,5,0)	' order of the info in the pattern (from 1 to 6): Y,M,D,H,M,S
	.PayeeLocation = 0		' start of payee in memo
	.PayeeLength = 0		' length of payee in memo
	.MonthNames = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", _
                       "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
' Trans Date,Description,Amount
	.Fields = Array( _
		Array(fldBookDate, "Trans Date"), _
		Array(fldMemo, "Description"), _
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
	Set .StatementCallback = GetRef("StatementCallback")
	Set .CustomDateCallback = GetRef("CustomDateCallback")
	Set .CustomAmountCallback = GetRef("CustomAmountCallback")
'	Set .ReadLineCallback = GetRef("ReadLineCallback")
End With

Dim aStates: Set aStates = CreateObject("Scripting.Dictionary")

Dim dStatementDate, dPrevBal, dClosingBal

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
    Dim sStates
    sStates = "ALAKAZARCACOCTDEFLGAHIIDILINIAKSKYLAMEMDMAMIMNMSMOMTNENVNHNJNMNYNCNDOHOKORPARISCSDTNTXUTVTVAWAWVWIWYASDCFMGUMHMPPWPRVIAEAAAP"
    Dim sState, i
    For i=1 To Len(sStates) Step 2
        sState = Mid(sStates, i, 2)
        aStates(sState) = 1
    Next
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
	LoadTextFile = DefaultLoadTextFile(Params)
End Function

' callback functions, called from DefaultRecogniseTextFile and DefaultLoadTextFile
' ths following implementations are functionally neutral or equivalent to the default processing
' in the class
Function HeaderCallback(sLine)
'MsgBox "In header callback: " & sLine
	Dim vFields
	Dim sTmp, iComma, sVal
	vFields = ParseLineDelimited(sLine, Params.CSVSeparator)
	If UBound(vFields) <> 2 Then
		HeaderCallback = True
		Exit Function
	End If
	Select Case vFields(1)
	Case "Statement Date:"
		dStatementDate = ParseDateEx(Trim(Replace(vFields(2), ",", "")), "MDY", " ")
msgbox "Statement date (" & vFields(2) & ") = " & dStatementDate
	Case "Account Number:"
		Params.AccountNum = Trim(Replace(vFields(2), " ", ""))
	Case "Previous Balance:"
		dPrevBal = ParseNumber(vFields(2), Params.DecimalSeparator)
	Case "Total Balance:"
		dClosingBal = ParseNumber(vFields(2), Params.DecimalSeparator)	
	End Select
	HeaderCallback = True
End Function

Function IsState(sWord)
    IsState = aStates.Exists(sWord)
End Function

Function TransactionCallback(t, vFields)
'MsgBox "In transaction callback: " & t.Memo
    Dim sMemo: sMemo = t.Memo
    Dim iLastSpace, sWord
    Dim iHash
' default to payee=memo
    t.Payee = sMemo
' if there is a store code, that's where we stop
    iHash = InStr(sMemo, "#")
    If iHash > 1 Then
        t.Payee = Trim(Left(sMemo, iHash-1))
    Else
' no store code, trim off "<city> <state>"
        iLastSpace = InStrRev(sMemo, " ")
        If iLastSpace > 0 Then
            sWord = Mid(sMemo, iLastSpace+1)
            If IsState(sWord) Then
                iLastSpace = InStrRev(sMemo, " ", iLastSpace-1)
                t.Payee = Trim(Left(sMemo, iLastSpace-1))
            End If
        End If
' lose any trailing digits
        t.Payee = TrimTrailingDigits(t.Payee)
    End If
    TransactionCallback = True
End Function

Function CustomDateCallback(sDate)
'MsgBox "In custom date callback: " & sDate
	Dim sTmp, dTmp
	sTmp = sDate & Left(Params.DateSeparator, 1) & Year(Now())
	dTmp = ParseDateEx(sTmp, Params.DateSequence, Params.DateSeparator)
	If dTmp <> NODATE Then
		dTmp = MostRecentEx(dTmp, dStatementDate)
    Else
        MsgBox ParseDateError
	End If
 msgbox "Date: " & sTmp & " is " & FormatDateTime(dTmp, 2)
	CustomDateCallback = dTmp
End Function
Function CustomAmountCallback(sAmt)
'MsgBox "In custom amount callback: " & sAmt
    Dim sTmp
    Dim iSign: iSign = 1
    If InStr(sAmt, "(") Then iSign = -1
    sTmp = Replace(sAmt, "(", "")
    sTmp = Replace(sTmp, ")", "")
    CustomAmountCallback = ParseNumber(sTmp, Params.DecimalSeparator) * iSign
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
Sub StatementCallback(stmt)
    stmt.ClosingBalance.BalDate = dStatementDate
    stmt.OpeningBalance.Amt = dPrevBal
    stmt.ClosingBalance.Amt = dClosingBal
End Sub
Function IsValidTxnLine(sLine)
'MsgBox "In IsValidTxnLine callback: " & sLine
	IsValidTxnLine = (IsNumeric(Mid(sLine, 2, 1)))
End Function
