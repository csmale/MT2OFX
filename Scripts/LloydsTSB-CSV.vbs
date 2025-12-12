' MT2OFX Input Processing Script LloydsTSB CSV format

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/LloydsTSB-CSV.vbs 6     15/06/09 19:25 Colin $"

Const ScriptName = "LloydsTSB-CSV"
Const FormatName = "LloydsTSB CSV Format"
Const ParseErrorMessage = "Cannot parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

Const DebugRecognition = False	' enables debug code in recognition
Const BankCode = "LOYDGB21"
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
Const CSVSeparator = ","
Const TxnLinePattern = ""
Const MinFieldsExpected = 8
Const MaxFieldsExpected = 8
Const DateSequence = "DMY"	' must be DMY, MDY, or YMD
Const DateSeparator = "/-"	' can be empty for dates in e.g. "yyyymmdd" format
Const InvertSign = False	' make credits into debits etc
Const CurrencyCode = "GBP"	' default if not specified in file
Const NoAvailableBalance = True		' True if file does not contain "Available Balance" information
Dim AccountNum: AccountNum = ""		' default if not specified in file
Dim BranchCode: BranchCode = ""		' default if not specified in file
Const SkipHeaderLines = 0	' number of lines to skip before the transaction data
Const ColumnHeadersPresent = False	' are the column headers in the file?
Const DecimalSeparator = "."	' as used in amounts
Const MemoChunkLength = 0	' if memo field consists of fixed length chunks
Const TxnDatePattern = ".*(\d\d)\.(\d\d)\.(\d\d)\ (\d\d)\.(\d\d)"	' pattern to find transaction date in the memo
Dim TxnDateSequence: TxnDateSequence = Array(3,2,1,4,5,0)	' order of the info in the pattern: Y,M,D,H,M,S
Const PayeeLocation = 0		' start of payee in memo
Const PayeeLength = 0		' length of payee in memo
Dim MonthNames					' month names in dates
MonthNames = Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
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
Const fldBranch = 18

' Declare fields in the order they appear in the file as an array of arrays. The inner arrays
' contain a field ID from the list above followed by the exact column header.
' LloydsTSB Sortcodes are in the range 30-39, according to Wikipedia...By checking for a 3 as the
' first digit we will (hopefully) eliminate false positives
Const csPatSortCode = "3\d-\d\d-\d\d"
Const csPatTransCode = "[A-Z][A-Z/][A-Z]?"
Const csPatTransCodeOptional = "([A-Z][A-Z/][A-Z]?)?"
Dim aFields1, aFields2, aFields
aFields1 = Array( _
	Array(fldBookDate, "Booking Date"), _
	Array(fldSkip, "Transaction Code", csPatTransCode), _
	Array(fldBranch, "Sort Code", csPatSortCode), _
	Array(fldAccountNum, "Account Number", "\d{6,8}"), _
	Array(fldMemo, "Payee"), _
	Array(fldAmtDebit, "Debit Amount"), _
	Array(fldAmtCredit, "Credit Amount"), _
	Array(fldClosingBal, "Balance") _
)
aFields2 = Array( _
	Array(fldBranch, "Sort Code", csPatSortCode), _
	Array(fldAccountNum, "Account Number", "\d{6,8}"), _
	Array(fldBookDate, "Booking Date"), _
	Array(fldSkip, "Transaction Code", csPatTransCodeOptional), _
	Array(fldMemo, "Payee"), _
	Array(fldAmtDebit, "Debit Amount"), _
	Array(fldAmtCredit, "Credit Amount"), _
	Array(fldClosingBal, "Balance") _
)
Dim iTxnCodeField

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
aPropertyList = Array()

Sub Initialise()
    LogProgress ScriptVersion, "Initialise"
	If Not CheckVersion() Then
		Abort
	End If
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

Sub SelectFormat(vFields)
' switch between old and new column orders
	If StringMatches(vFields(1), csPatSortCode) Then	' is it a sort code?
		aFields = aFields2
		iTxnCodeField = 4
	Else
		aFields = aFields1
		iTxnCodeField = 2
	End If
End Sub

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
	SelectFormat vFields
	If ColumnHeadersPresent Then
		For i=1 To UBound(vFields)
			If vFields(i) <> aFields(i-1)(1) Then
				If DebugRecognition Then
					MsgBox "Field " & CStr(i) & " " & aFields(i-1)(1) & " instead of " & vFields(i),,ScriptName
				End If
				Exit function
			End If
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
				Case fldBranch
					sPat = "(branch code)"
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
	Dim sBranch		' branch code
	Dim Stmt        ' holds the current statement
	Dim sTmp		' temporary string
	Dim vDateBits	' parts of date
	Dim vTmp		' temp for splitting memo
	Dim iSeq		' transaction sequence number
	Dim iTmp
	Dim i
	Dim dBal		' temp balance date
	Dim sField		' field value being processed
	Dim dMaxDate	' latest txn/book date - if we don't have a statement date
	Dim dMinDate: dMinDate = NODATE: dMaxDate = NODATE
	Dim daStatement: Set daStatement = New DateAccumulator

	LoadTextFile = False
	sAcct = ""
	For i=1 To SkipHeaderLines
		sLine = ReadLine()
	Next
	If ColumnHeadersPresent Then
		sLine = ReadLine()
	End If
	' fill field lookup dictionary
' NB: only the last occurrence is remembered!
' if we have been through recognition, set up FieldDict now. Otherwise it will get sorted on the
' first transaction.
	If Not IsEmpty(aFields) Then
		For i=0 To UBound(aFields)
			FieldDict(aFields(i)(0)) = i+1
		Next
	End If

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
			If IsEmpty(aFields) Then
				SelectFormat vFields
	' fill field lookup dictionary
' NB: only the last occurrence is remembered!
				For i=0 To UBound(aFields)
					FieldDict(aFields(i)(0)) = i+1
				Next
			End If
	' set up new transaction, and start a new statement if the account # changes
			If sAcct <> vFields(FieldDict(fldAccountNum)) Then
				Set Stmt = NewStatement()
		' this initialisation should be in the class constructor!! (fixed in 3.3.5)
				Stmt.OpeningBalance.BalDate = NODATE
				Stmt.OpeningBalance.Ccy = CurrencyCode
				Stmt.AvailableBalance.BalDate = NODATE
				If Not NoAvailableBalance Then Stmt.AvailableBalance.Ccy = CurrencyCode
				Stmt.ClosingBalance.BalDate = NODATE				
				Stmt.ClosingBalance.Ccy = CurrencyCode
				iSeq = 0
				Stmt.BankName = BankCode
				Stmt.BranchName = BranchCode
				dMaxDate = NODATE
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
				case fldBranch
					Stmt.BranchName = sField
					sBranch = sField
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
					daStatement.Process Txn.BookDate
				case fldValueDate
					Txn.ValueDate = ParseDate(sField)
				Case fldTransactionDate
					Txn.TxnDate = ParseDate(sField)
					Txn.TxnDateValid = (Txn.TxnDate <> NODATE)
					daStatement.Process Txn.TxnDate
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
					If dBal > Stmt.ClosingBalance.BalDate Or Stmt.ClosingBalance.BalDate = NODATE Then
						Stmt.ClosingBalance.BalDate = dBal
						Stmt.AvailableBalance.BalDate = dBal
					End If
					If dBal < Stmt.OpeningBalance.BalDate Or Stmt.OpeningBalance.BalDate = NODATE Then
						Stmt.OpeningBalance.BalDate = dBal
					End If
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
			
			Dim sMemo
' find the payee, transaction type and txn date if we can
			sMemo = Txn.Memo
			If Left(sMemo, 4) = "FPC/" Then
				sTmp = Replace(sMemo, " . ", vbTab)
				vTmp = Split(sTmp, vbTab)
				Txn.Payee = Trim(vTmp(2))
			Else
				Txn.Payee = Trim(Left(sMemo, 18))
				iTmp = InStr(Txn.Payee, " . ")
				If iTmp > 0 Then
					Txn.Payee = Trim(Left(Txn.Payee, iTmp-1))
				End If
			End If
			Select Case vFields(iTxnCodeField)
			Case "CHQ"
				Txn.TxnType = "CHECK"
				Txn.CheckNum = Txn.Payee
				Txn.Payee = ""
			Case "DD", "D/D"
				Txn.TxnType = "DIRECTDEBIT"
			Case "SO", "S/O"
				If Txn.Amt < 0 Then
					Txn.TxnType = "REPEATPMT"
				End If
			Case "FPI"	' FastPay?
				Txn.TxnType = "PAYMENT"
			Case "BGC"
				Txn.TxnType = "CREDIT"
			Case "TFR"
				Txn.TxnType = "XFER"
			Case "CHG"
				Txn.TxnType = "SRVCHG"
			Case "PAY"
				Txn.TxnType = "CHECK"
				Txn.CheckNum = Txn.Payee
				Txn.Payee = ""
			Case "CPT"
				Txn.TxnType = "ATM"
				Txn.Payee = "Cash Withdrawal"
			End Select
						
' tidy up the memo
			If MemoChunkLength > 0 Then
				sMemo = Txn.Memo
				Txn.Memo = ""
				For i=1 To Len(sMemo) Step MemoChunkLength
					ConcatMemo Trim(Mid(sMemo, i, MemoChunkLength))
				Next
			End If

' keep tabs on the statement/balance Date
			Stmt.ClosingBalance.BalDate = daStatement.MaxDate
			Stmt.OpeningBalance.BalDate = daStatement.MinDate
		End If
	Loop
	LoadTextFile = True
End Function

Private Function TransDate(sMemo)
	Dim vDateBits
	Dim dTxn
	Dim iYear, iMonth, iDay, iHour, iMin, iSec
	dTxn = NODATE
	vDateBits = ParseLineFixed(sMemo, TxnDatePattern)
	If TypeName(vDateBits) = "Variant()" Then
		If TxnDateSequence(0) > 0 Then iYear = CInt(vDateBits(TxnDateSequence(0)))
		If TxnDateSequence(1) > 0 Then iMonth = CInt(vDateBits(TxnDateSequence(1)))
		If TxnDateSequence(2) > 0 Then iDay = CInt(vDateBits(TxnDateSequence(2)))
		If TxnDateSequence(3) > 0 Then iHour = CInt(vDateBits(TxnDateSequence(3)))
		If TxnDateSequence(4) > 0 Then iMin = CInt(vDateBits(TxnDateSequence(4)))
		If TxnDateSequence(5) > 0 Then iSec = CInt(vDateBits(TxnDateSequence(5)))
		dTxn = DateSerial(iYear, iMonth, iDay) + TimeSerial(iHour, iMin, iSec)
	End If
	TransDate = dTxn
End Function
