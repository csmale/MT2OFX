' MT2OFX Input Processing Script Basic CSV format
' NB: This Script Will Not Work Without Customisation!

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/AbbeyUK2-CSV.vbs 3     4/12/09 22:10 Colin $"

Dim Params: Set Params = New MT2OFXScript

With Params
	.MinimumProgramVersion = "3"
	.DebugRecognition = False	' enables debug code in recognition
	.ScriptName = "AbbeyUK2-CSV"
	.FormatName = "Abbey Bank (UK) Consumer CSV"
	.ParseErrorMessage = "Cannot parse line."
	.ParseErrorTitle = .ScriptName
	.BankCode = "ABBYGB2L"
	.AccountNum = ""		' default if not specified in file
	.BranchCode = ""		' default if not specified in file
	.QuickenBankID = "10058"
	.CurrencyCode = "GBP"	' default if not specified in file
	.ColumnHeadersPresent = False	' are the column headers in the file?
	.SkipHeaderLines = 0	' number of lines to skip before the transaction data
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
	.CSVSeparator = ","
	.TxnLinePattern = ""
	.DecimalSeparator = "."	' as used in amounts
	.DateSequence = "DMY"	' must be DMY, MDY, or YMD
	.DateSeparator = "-/. "	' can be empty for dates in e.g. "yyyymmdd" format
	.OldestLast = True		' True if transactions are in reverse order
	.InvertSign = False	' make credits into debits etc
	.NoAvailableBalance = True		' True if file does not contain "Available Balance" information
	.MemoChunkLength = 0	' if memo field consists of fixed length chunks
	.TxnDatePattern = ".* ON (\d\d\d\d)-(\d\d)-(\d\d)"	' pattern to find transaction date in the memo
	.TxnDateSequence = Array(1,2,3,0,0,0)	' order of the info in the pattern (from 1 to 6): Y,M,D,H,M,S
	.PayeeLocation = 0		' start of payee in memo
	.PayeeLength = 0		' length of payee in memo
	.MonthNames = Empty
' 111111 22222222,15/08/2008,CARD PAYMENT TO BLABLA ON 2008-08-12,-83.40
' first field has sort code followed by account number. these are sorted out in the FinaliseCallback
	.Fields = Array( _
		Array(fldAccountNum, "Account number", "\d{6} \d{8}"), _
		Array(fldBookDate, "Booking date"), _
		Array(fldMemo, "Memo"), _
		Array(fldAmount, "Amount") _
	)
' min/max fields expected: default to size of Fields array. can be overridden here if required
'	.MinFieldsExpected = 1
'	.MaxFieldsExpected = 1
	.Properties = Array()

' callbacks from standard code in MT2OFX.vbs back to this script
	Set .TransactionCallback = GetRef("TransactionCallback")
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

Function TransactionCallback(t, vFields)
	Dim iEnd: iEnd = 0
	Dim iTmp, sTmp, sMemo
    Dim re, m, dTxn
    Set re=New RegExp
' MsgBox "In transaction callback: " & t.Memo
' get transaction date
    sMemo = t.Memo
    dTxn = NODATE
    re.Pattern = ".* ON (\d\d\d\d)-(\d\d)-(\d\d)"
    Set m = re.Execute(sMemo)
    If m.Count > 0 Then
        dTxn = DateSerial(CInt(m(0).SubMatches(0)), CInt(m(0).SubMatches(1)), CInt(m(0).SubMatches(2)))
        sTmp = Left(sMemo, Len(sMemo) - 14)
    Else
        re.Pattern = ".* ON (\d\d)-(\d\d)-(\d\d\d\d)"
        Set m = re.Execute(sMemo)
        If m.Count > 0 Then
            dTxn = DateSerial(CInt(m(0).SubMatches(2)), CInt(m(0).SubMatches(1)), CInt(m(0).SubMatches(0)))
            sTmp = Left(sMemo, Len(sMemo) - 14)
        End If
    End If
    If dTxn <> NODATE Then
        t.TxnDateValid = True
        t.TxnDate = dTxn
        sMemo = sTmp
    End If
' get transaction amount (may be in foreign currency?)
    re.Pattern = ".* (\d+\.\d\d [A-Z][A-Z][A-Z])"
    Set m=re.Execute(sMemo)
    If m.Count > 0 Then
        sMemo = Left(sMemo, Len(sMemo) - Len(m(0).SubMatches(0)) - 1)
    End If

    If StartsWith(sMemo, "CARD PAYMENT TO ") Then
        t.TxnType = "POS"
        t.Payee = Mid(sMemo, 17)
        iEnd = InStr(t.Payee, " ON ")
    ElseIf StartsWith(sMemo, "BILL PAYMENT TO ") Then
        t.Payee = Mid(sMemo, 17)
        iEnd = InStr(t.Payee, " REFERENCE ")
    ElseIf StartsWith(sMemo, "DIRECT DEBIT PAYMENT TO ") Then
        t.TxnType = "DIRECTDEBIT"
        t.Payee = Mid(sMemo, 25)
        iEnd = InStr(t.Payee, " REF ")
    ElseIf StartsWith(sMemo, "PAYMENT MADE BY CHEQUE") Or StartsWith(t.Memo, "PAYMENT BY CHEQUE") Then
        t.TxnType = "CHECK"
        iTmp = InStr(sMemo, "SERIAL NO ")
        If iTmp>0 Then
            t.CheckNum = Mid(sMemo, iTmp+10)
        End If
    ElseIf StartsWith(sMemo, "INTEREST") Then
        t.TxnType = "INT"
    ElseIf StartsWith(sMemo, "CASH OUT AT ") Then
        t.TxnType = "ATM"
    ElseIf StartsWith(sMemo, "REFUND OF ACCOUNT FEE") Then
        t.TxnType = "FEE"
    ElseIf StartsWith(sMemo, "UNPAID DIRECT DEBIT FEE") Then
        t.TxnType = "FEE"
    ElseIf StartsWith(sMemo, "UNPAID CHEQUE") Then
        t.TxnType = "FEE"
    ElseIf StartsWith(sMemo, "UNPAID DIRECT DEBIT") Then
        t.TxnType = "FEE"
    ElseIf StartsWith(sMemo, "PAID ITEM FEE") Then
        t.TxnType = "FEE"
    ElseIf StartsWith(sMemo, "UNAUTHORISED OVERDRAFT FEE") Then
        t.TxnType = "FEE"
    End If
    If iEnd<> 0 Then
        t.Payee = Left(t.Payee, iEnd-1)
    End If
	TransactionCallback = True
End Function
Function FinaliseCallback()
' MsgBox "In finalisation callback"
	Dim xStmt, sTmp
	For Each xStmt In Session.Statements
		sTmp = xStmt.Acct
		xStmt.Acct = Mid(sTmp, 8)
		xStmt.BranchName = Left(sTmp, 6)
	Next
	FinaliseCallback = True
End Function
