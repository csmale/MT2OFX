' MT2OFX Input Processing Script Co-op Bank (GB) CSV format

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/CoopGB-CSV.vbs 4     29/01/10 22:00 Colin $"

Const ScriptName = "CoopGB-CSV"
Const FormatName = "Co-op Bank (GB) CSV"
Const ParseErrorMessage = "Cannot parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

Const DebugRecognition = False	' enables debug code in recognition
Const BankCode = "CPBKGB22"
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
Const CSVSeparator = ","
Const TxnLinePattern = ""
Const MinFieldsExpected = 8
Const MaxFieldsExpected = 9
Const DateSequence = "DMY"	' must be DMY, MDY, or YMD
Const DateSeparator = "/"	' can be empty for dates in e.g. "yyyymmdd" format
Const InvertSign = False	' make credits into debits etc
Const CurrencyCode = "GBP"	' default if not specified in file
Const AccountNum = ""		' default if not specified in file
Const SkipHeaderLines = 0	' number of lines to skip before the transaction data
Const ColumnHeadersPresent = True	' are the column headers in the file?
Const DecimalSeparator = "."	' as used in amounts
Const MemoChunkLength = 0	' if memo field consists of fixed length chunks
Const TxnDatePattern = ".*(\d\d)\.(\d\d)\.(\d\d)\ (\d\d)\.(\d\d)"	' pattern to find transaction date in the memo
Const PayeeLocation = 0		' start of payee in memo
Const PayeeLength = 0		' length of payee in memo
Dim MonthNames					' month names in dates
'MonthNames = Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
' Either give the month names in an array as above or use SetLocale to get the
' system strings for the given locale. Otherwise the default locale will be used.
' The MonthNames array must have a multiple of 12 elements, which run from Jan-Dec in groups of
' 12, i.e. "Jan".."Dec","January".."December" etc. Lower/upper case is not significant.
' SetLocale "nl-nl"

Const fldSkip = 0
Const fldAccountNum = 1
Const fldCurrency = 2
Const fldClosingBal = 3
Const fldAvailBal = 4
Const fldBookDate = 5
Const fldValueDate = 6
Const fldAmtCredit = 7
Const fldAmtDebit = 8
Const fldMemo = 9
Const fldBalanceDate = 10
Const fldAmount = 11
Const fldPayee = 12
Const fldTransactionDate = 13
Const fldTransactionTime = 14
Const fldChequeNum = 15
Const fldCheckNum = 15
Const fldFITID = 16
Const fldEmpty = 17	' field is ignored but MUST be empty for recognition

' Declare fields in the order they appear in the file as an array of arrays. The inner arrays
' contain a field ID from the list above followed by the exact column header.
' Trans Date,Description,Bank Ref,Customer Ref,Credit,Debit,Running Balance,Additional Information
Dim aFields
aFields = Array( _
	Array(fldBookDate, "Trans Date"), _
	Array(fldMemo, "Description"), _
	Array(fldMemo, "Bank Ref"), _
	Array(fldMemo, "Customer Ref"), _
	Array(fldAmtCredit, "Credit"), _
	Array(fldAmtDebit, "Debit"), _
	Array(fldClosingBal, "Running Balance"), _
	Array(fldMemo, "Additional Information"), _
	Array(fldSkip, "") _
)

' Dictionary to facilitate field lookup by field code
Dim FieldDict
Set FieldDict = CreateObject("Scripting.Dictionary")

' Property List is an array of arrays, each of which has the following elements:
'	1. Property key - used to reference properties
'	2. Property name - used as a label in the config screen
'	3. Property description - used as a description or tooltip in the config screen
'	4. Data type - ptString, ptBoolean, ptInteger, ptFloat, ptDate, ptChoose
'	5. Value list (will be displayed in a combobox) - array of values (Only with ptChoose)
Dim aPropertyList
aPropertyList = Array( _
	Array("AcctNum", "Account Number", _
		"Please enter the account number.", _
		ptString,,"=ValidAccount") _
	)
Private xValidationMessage
Function ValidAccount(s)
	If Len(s) = 8 Then
		ValidAccount = True
	Else
		ValidAccount = False
		xValidationMessage = "Bank account number must be 8 digits. '%1' is wrong."
	End If
End Function
Function ValidationMessage
	ValidationMessage = xValidationMessage
End Function

Sub Initialise()
    LogProgress ScriptVersion, "Initialise"
	If Not CheckVersion() Then
		Abort
	End If
' fill field lookup dictionary
' NB: only the last occurrence is remembered!
	Dim i
	For i=0 To UBound(aFields)
		FieldDict(aFields(i)(0)) = i+1
	Next
' Initialise dictionary of month names
	InitialiseMonths MonthNames
' get properties
	LoadProperties ScriptName, aPropertyList
End Sub

' function DescriptiveName
' returns a string with a descriptive name of this script
Function DescriptiveName()
	DescriptiveName = FormatName
End Function

Sub Configure
	If ShowConfigDialog(ScriptName, aPropertyList) Then
		SaveProperties ScriptName, aPropertyList
	End If
End Sub

Function ParseDate(sDate)
	ParseDate = ParseDateEx(sDate, DateSequence, DateSeparator)
End Function

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
	Dim vFields
	Dim sLine
	Dim i
	Dim bTmp
	Dim sField
	Dim sPat
	RecogniseTextFile = False
	For i=1 To SkipHeaderLines
		If AtEOF() Then
			Exit Function
		End If
		sLine = ReadLine()
	Next
	If AtEOF() Then
		Exit Function
	End If
	sLine = ReadLine()
	If CSVSeparator = "" Then
		vFields = ParseLineFixed(sLine, TxnLinePattern)
	Else
		vFields = ParseLineDelimited(sLine, CSVSeparator)
	End If
	If TypeName(vFields) <> "Variant()" Then
		If DebugRecognition Then
			MsgBox "not var array",,ScriptName
		End If
		Exit Function
	End If
	If UBound(vFields) < MinFieldsExpected Or UBound(vFields) > MaxFieldsExpected Then
		If DebugRecognition Then
			MsgBox "Wrong number of fields - got " & UBound(vFields) & ", expected " _
			& MinFieldsExpected & "-" & MaxFieldsExpected & " - " & sLine,,ScriptName
		End If
		Exit function
	End If
	If ColumnHeadersPresent Then
		For i=1 To UBound(vFields)
			If vFields(i) <> aFields(i-1)(1) Then
				If DebugRecognition Then
					MsgBox "Field " & CStr(i) & " " & aFields(i-1)(1) & " instead of " & vFields(i),,ScriptName
				End If
				Exit function
			End if
		Next
	Else
' pattern-match the first row
		For i=1 To UBound(vFields)
			sField = Trim(vFields(i))
			If UBound(aFields(i-1)) > 1 Then
				sPat = aFields(i-1)(2)
				bTmp = StringMatches(sField, sPat)
			Else
				Select Case aFields(i-1)(0)
				case fldSkip, fldMemo, fldPayee
					bTmp = True
				Case fldEmpty
					sPat = "(empty)"
					bTmp = (Len(sField) = 0)
				Case fldAccountNum
					sPat = "(account number)"
					bTmp = (Len(sField) > 0)
				case fldCurrency
					sPat = "[A-Z][A-Z][A-Z]"
					bTmp = StringMatches(sField, sPat)
				case fldClosingBal, fldAvailBal, fldAmtCredit, fldAmtDebit, fldAmount
					If DecimalSeparator = "." Then
						sPat = "[+-]?[ 0-9,]*(\.[0-9]*)?"
					Else
						sPat = "[+-]?[ 0-9\.]*(,[0-9]*)?"
					End If
					bTmp = StringMatches(sField, sPat)
				case fldBookDate, fldValueDate, fldTransactionDate, fldBalanceDate
	' NB: ParseDate will throw an error on an invalid date! need to sort this
					sPat = "(date)"
					bTmp = (ParseDate(sField) <> NODATE)
				Case fldTransactionTime
					sPat = "(time)"
					bTmp = (Len(sField) > 0)
				End Select
			End If
			If Not bTmp Then
				If DebugRecognition Then
					MsgBox "Field " & i & " (" & sField & ") failed to match '" & sPat & "'",,ScriptName
				End If
				Exit Function
			End If
		Next
	End If
	LogProgress ScriptName, "File Recognised"
	RecogniseTextFile = True
End Function

Function LoadTextFile()
	Dim sLine       ' holds a line
	Dim vFields     ' array of fields in the line
	Dim sAcct       ' last account number
	Dim Stmt        ' holds the current statement
	Dim sTmp		' temporary string
	Dim vDateBits	' parts of date
	Dim iSeq		' transaction sequence number
	Dim i
	Dim dBal		' temp balance date
	Dim sField		' field value being processed
   Dim da: Set da = New DateAccumulator

	LoadTextFile = False
	sAcct = GetProperty("AcctNum")
	If Len(sAcct) = 0 Then
		Message True,True, "Please configure the account number first: Options->Scripts->Parameters", ScriptName
		Exit Function
	End If
	For i=1 To SkipHeaderLines
		sLine = ReadLine()
	Next
	If ColumnHeadersPresent Then
		sLine = ReadLine()
	End if
	Do While Not AtEOF()
		sLine = ReadLine()
		If Len(sLine) > 0 And Left(sLine,1) <> CSVSeparator Then
			If CSVSeparator = "" Then
				vFields = ParseLineFixed(sLine, TxnLinePattern)
			Else
				vFields = ParseLineDelimited(sLine, CSVSeparator)
			End If
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			If UBound(vFields) < MinFieldsExpected Or UBound(vFields) > MaxFieldsExpected Then
				Message True, True, "Wrong number of fields - " & CStr(UBound(vFields)+1) & " - " & sLine, ScriptName
				Abort
				Exit function
			End If
	' set up new transaction, and start a new statement if the account # changes
			If FieldDict.Exists(fldAccountNum) Then
				If sAcct <> vFields(FieldDict(fldAccountNum)) Then
                da.Reset
					Set Stmt = NewStatement()
		' this initialisation should be in the class constructor!! (fixed in 3.3.5)
					Stmt.OpeningBalance.BalDate = NODATE
					Stmt.OpeningBalance.Ccy = CurrencyCode
					Stmt.AvailableBalance.BalDate = NODATE
					Stmt.AvailableBalance.Ccy = ""
					Stmt.ClosingBalance.BalDate = NODATE				
					Stmt.ClosingBalance.Ccy = CurrencyCode
					iSeq = 0
					Stmt.BankName = BankCode
				End If
			Else
				If IsEmpty(Stmt) Then
                da.Reset
					Set Stmt = NewStatement()
		' this initialisation should be in the class constructor!! (fixed in 3.3.5)
					Stmt.OpeningBalance.BalDate = NODATE
					Stmt.OpeningBalance.Ccy = CurrencyCode
					Stmt.AvailableBalance.BalDate = NODATE
					Stmt.AvailableBalance.Ccy = ""
					Stmt.ClosingBalance.BalDate = NODATE				
					Stmt.ClosingBalance.Ccy = CurrencyCode
					iSeq = 0
					Stmt.BankName = BankCode
					Stmt.Acct = sAcct
				End If
			End If
			NewTransaction
			iSeq = iSeq + 1
			LastMemo = ""
			For i=1 To UBound(vFields)
				sField = Trim(vFields(i))
				Select Case aFields(i-1)(0)
				case fldSkip, fldEmpty
					' do nothing
				case fldAccountNum
					Stmt.Acct = sField
					sAcct = sField
				case fldCurrency
					Stmt.OpeningBalance.Ccy = sField
					Stmt.ClosingBalance.Ccy = sField
					Stmt.AvailableBalance.Ccy = sField
				case fldClosingBal
					Stmt.ClosingBalance.Amt = ParseNumber(sField, DecimalSeparator)
				case fldAvailBal
					Stmt.AvailableBalance.Amt = ParseNumber(sField, DecimalSeparator)
				case fldBookDate
					Txn.BookDate = ParseDate(sField)
               da.Process Txn.BookDate
				case fldValueDate
					Txn.ValueDate = ParseDate(sField)
				Case fldTransactionDate
					Txn.TxnDate = ParseDate(sField)
               da.Process Txn.TxnDate
				Case fldTransactionTime
					If Txn.TxnDate <> NODATE And Len(sField)=5 Then
						Txn.TxnDate = Txn.TxnDate + TimeSerial(CInt(Left(sField,2)), _
							CInt(Mid(sField,4,2)),0)
					End If
				case fldAmtCredit
					Txn.Amt = Txn.Amt + Abs(ParseNumber(sField, DecimalSeparator))
				case fldAmtDebit
					Txn.Amt = Txn.Amt - Abs(ParseNumber(sField, DecimalSeparator))
				Case fldAmount
					Txn.Amt = ParseNumber(sField, DecimalSeparator)
				Case fldChequeNum
					Txn.CheckNum = sField
				case fldMemo
					ConcatMemo sField
				Case fldBalanceDate
					dBal = ParseDate(sField)
               da.Process dBal
				Case fldPayee
					If Len(sField) > 0 Then
						Txn.Payee = sField
					End If
				Case fldFITID
					Txn.FITID = sField
				End select
			Next
' correct the sign of the amount
			If InvertSign Then
				Txn.Amt = -Txn.Amt
			End If

' transaction type
			If Txn.Amt < 0 Then
				Txn.TxnType = "PAYMENT"
			Else
				Txn.TxnType = "DEP"
			End If

' Co-op GB
            If vFields(2) = "Cheque" Then
                Txn.TxnType = "CHECK"
                Txn.CheckNum = vFields(4)
            ElseIf vFields(2) = "Direct Debit" Then
                Txn.TxnType = "DIRECTDEBIT"
                Txn.Payee = vFields(4)
            ElseIf vFields(2) = "First Direct Debit" Then
                Txn.TxnType = "DIRECTDEBIT"
                Txn.Payee = vFields(4)
            ElseIf vFields(2) = "Faster Payment" Then
                Txn.Payee = vFields(4)
            ElseIf vFields(2) = "Purchase" Then
                Txn.TxnType = "POS"
                Txn.Payee = Trim(Mid(vFields(4), 5))
            ElseIf vFields(2) = "Refund" Then
                Txn.TxnType = "POS"
                Txn.Payee = Trim(Mid(vFields(4), 5))
            ElseIf vFields(2) = "Standing Order" Then
                Txn.Payee = vFields(4)
            ElseIf vFields(2) = "BACS Credit" Then
                Txn.Payee = vFields(4)
            ElseIf vFields(2) = "ATM Debit" Then
                Txn.TxnType = "ATM"
            ElseIf vFields(2) = "Bank Credit Interest" Then
                Txn.TxnType = "INT"
            End If
            
            Stmt.ClosingBalance.BalDate = da.MaxDate
            Stmt.OpeningBalance.BalDate = da.MinDate
			
			Dim sMemo
' find the payee, transaction type and txn date if we can
			sMemo = Txn.Memo
			If PayeeLocation > 0 And Len(Txn.Payee) = 0 Then
				Txn.Payee = Trim(Mid(sMemo, PayeeLocation, PayeeLength))
			End If
			If Len(TxnDatePattern) > 0 Then
				vDateBits = ParseLineFixed(Txn.Memo, TxnDatePattern)
				If TypeName(vDateBits) = "Variant()" Then
                If UBound(vDateBits) = 4 Then
                    Txn.TxnDate = DateSerial(Year(Stmt.OpeningBalance.BalDate), CInt(vDateBits(2)), CInt(vDateBits(1))) _
                    + TimeSerial(CInt(vDateBits(3)), CInt(vDateBits(4)), 0)
                    Txn.TxnDateValid = True
                End If
				End If
			End If

' tidy up the memo
			If MemoChunkLength > 0 Then
				sMemo = Txn.Memo
				Txn.Memo = ""
				For i=1 To Len(sMemo) Step MemoChunkLength
					ConcatMemo Trim(Mid(sMemo, i, MemoChunkLength))
				Next
			End If

' keep tabs on the statement/balance Date
			If Stmt.ClosingBalance.BalDate = NODATE Then
				Stmt.ClosingBalance.BalDate = dMaxDate
			End If
		End If
	Loop
	LoadTextFile = True
End Function
