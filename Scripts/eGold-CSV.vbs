' MT2OFX Input Processing Script for eGold CSV format

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/eGold-CSV.vbs 3     2/11/05 23:03 Colin $"

Const ScriptName = "eGold-CSV"
Const FormatName = "eGold CSV format"
Const ParseErrorMessage = "Cannot parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

Const DebugRecognition = False	' enables debug code in recognition
Const BankCode = "EGOLD"
Const CSVSeparator = ","
Const NumFieldsExpected = 15
Const DateSequence = "MDY"	' must be DMY, MDY, or YMD
Const DateSeparator = "/"	' can be empty for dates in e.g. "yyyymmdd" format
Const CurrencyCode = "XAU"	' default if not specified in file
Const AccountNum = "12345"	' if not specified in the file
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

' Declare fields in the order they appear in the file as an array of arrays. The inner arrays
' contain a field ID from the list above followed by the exact column header.
'Time,Type,Batch,e-metal,Weight,To/From Number,To/From Name,
'Entered Amount,Entered Currency,Rate,Memo,PayeeName,PayeeAddress,Redemption Info,Merchant Ref Number
Dim aFieldsXAU
aFieldsXAU = Array( _
	Array(fldBookDate, "Time"), _
	Array(fldSkip, "Type"), _
	Array(fldChequeNum, "Batch"), _
	Array(fldSkip, "e-metal"), _
	Array(fldAmount, "Weight"), _
	Array(fldSkip, "To/From Number"), _
	Array(fldPayee, "To/From Name"), _
	Array(fldSkip, "Entered Amount"), _
	Array(fldSkip, "Entered Currency"), _
	Array(fldSkip, "Rate"), _
	Array(fldMemo, "Memo"), _
	Array(fldSkip, "PayeeName"), _
	Array(fldSkip, "PayeeAddress"), _
	Array(fldSkip, "Redemption Info"), _
	Array(fldSkip, "Merchant Ref Number") _
)

' This version accounts for the real money and not the gold. All transactions are assumed To
' be in the same (real) currency, which will override XAU!
Dim aFieldsXXX
aFieldsXXX = Array( _
	Array(fldBookDate, "Time"), _
	Array(fldSkip, "Type"), _
	Array(fldChequeNum, "Batch"), _
	Array(fldSkip, "e-metal"), _
	Array(fldSkip, "Weight"), _
	Array(fldSkip, "To/From Number"), _
	Array(fldPayee, "To/From Name"), _
	Array(fldAmount, "Entered Amount"), _
	Array(fldCurrency, "Entered Currency"), _
	Array(fldSkip, "Rate"), _
	Array(fldMemo, "Memo"), _
	Array(fldSkip, "PayeeName"), _
	Array(fldSkip, "PayeeAddress"), _
	Array(fldSkip, "Redemption Info"), _
	Array(fldSkip, "Merchant Ref Number") _
)

' aFields gets set to either aFieldsXXX or aFieldsXAU according to the script property eGoldUseRealCurrency
Dim aFields

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
	Array("eGoldUseRealCurrency", "Use Real Currency", _
		"If this value is True, transactions will be recorded in the real currency." & _
		" Otherwise the quantity of gold will be used as the currency XAU.", _
		ptBoolean) _
	)

Sub Initialise()
    LogProgress ScriptVersion, "Initialise"
	If Not CheckVersion() Then
		Abort
	End If
' get properties - must do this early for eGold!
	LoadProperties ScriptName, aPropertyList' fill field lookup dictionary
	If CBool(GetProperty("eGoldUseRealCurrency")) Then
		aFields = aFieldsXXX
	Else
		aFields = aFieldsXAU
	End If
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
	vFields = ParseLineDelimited(sLine, CSVSeparator)
	If TypeName(vFields) <> "Variant()" Then
		If DebugRecognition Then
			MsgBox "not var array"
		End If
		Exit Function
	End If
	If UBound(vFields) <> NumFieldsExpected Then
		If DebugRecognition Then
			MsgBox "wrong number of fields - got " & UBound(vFields) & ", expected " _
			& NumFieldsExpected & " - " & sLine
		End If
		Exit function
	End If
	If ColumnHeadersPresent Then
		For i=1 To NumFieldsExpected
			If vFields(i) <> aFields(i-1)(1) Then
				If DebugRecognition Then
					MsgBox "field " & CStr(i) & " " & aFields(i-1)(1) & " instead of " & vFields(i)
				End If
				Exit function
			End if
		Next
	Else
' pattern-match the first row
		For i=1 To NumFieldsExpected
			sField = Trim(vFields(i))
			If UBound(aFields(i-1)) > 2 Then
				bTmp = StringMatches(sField, aFields(i-1)(2))
			Else
				Select Case aFields(i-1)(0)
				case fldSkip, fldMemo, fldPayee
					bTmp = True
				Case fldAccountNum
					bTmp = (Len(sField) > 0)
				case fldCurrency
					bTmp = StringMatches(s, "[A-Z][A-Z][A-Z]")
				case fldClosingBal, fldAvailBal, fldAmtCredit, fldAmtDebit, fldAmount
					If DecimalSeparator = "." Then
						sPat = "[ 0-9,]*(\.[0-9]*)?"
					Else
						sPat = "[ 0-9\.]*(,[0-9]*)?"
					End If
					bTmp = StringMatches(sField, sPat)
				case fldBookDate, fldValueDate, fldTransactionDate, fldBalanceDate
	' NB: ParseDate will throw an error on an invalid date! need to sort this
					bTmp = (ParseDate(sField) <> NODATE)
				Case fldTransactionTime
					bTmp = (Len(sField) > 0)
				End Select
			End If
			If Not bTmp Then
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

	LoadTextFile = False
	sAcct = ""
	For i=1 To SkipHeaderLines
		sLine = ReadLine()
	Next
	If ColumnHeadersPresent Then
		sLine = ReadLine()
	End if
	Do While Not AtEOF()
		sLine = ReadLine()
		If Len(sLine) > 0 And Left(sLine,1) <> CSVSeparator Then
			vFields = ParseLineDelimited(sLine, CSVSeparator)
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			If UBound(vFields) > NumFieldsExpected Or UBound(vFields) < (NumFieldsExpected-5) Then
				Message True, True, "Wrong number of fields - " & CStr(UBound(vFields)+1) & " - " & sLine, ScriptName
				Abort
				Exit function
			End If
	' set up new transaction, and start a new statement if the account # changes
			If FieldDict.Exists(fldAccountNum) Then
				If sAcct <> vFields(FieldDict(fldAccountNum)) Then
					Set Stmt = NewStatement()
		' this initialisation should be in the class constructor!! (fixed in 3.3.5)
					Stmt.OpeningBalance.BalDate = NODATE
					Stmt.OpeningBalance.Ccy = CurrencyCode
					Stmt.AvailableBalance.BalDate = NODATE
					Stmt.AvailableBalance.Ccy = CurrencyCode
					Stmt.ClosingBalance.BalDate = NODATE				
					Stmt.ClosingBalance.Ccy = CurrencyCode
					iSeq = 0
					Stmt.BankName = BankCode
				End If
			Else
				If IsEmpty(Stmt) Then
					Set Stmt = NewStatement()
		' this initialisation should be in the class constructor!! (fixed in 3.3.5)
					Stmt.OpeningBalance.BalDate = NODATE
					Stmt.OpeningBalance.Ccy = CurrencyCode
					Stmt.AvailableBalance.BalDate = NODATE
					Stmt.AvailableBalance.Ccy = CurrencyCode
					Stmt.ClosingBalance.BalDate = NODATE				
					Stmt.ClosingBalance.Ccy = CurrencyCode
					iSeq = 0
					Stmt.BankName = BankCode
					Stmt.Acct = AccountNum
				End If
			End If
			NewTransaction
			iSeq = iSeq + 1
			LastMemo = ""
			For i=1 To UBound(vFields)
				sField = Trim(vFields(i))
				Select Case aFields(i-1)(0)
				case fldSkip
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
				case fldValueDate
					Txn.ValueDate = ParseDate(sField)
				Case fldTransactionDate
					Txn.TxnDate = ParseDate(sField)
					Txn.TxnDateValid = (Txn.TxnDate <> NODATE)
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
				Case fldChequeNum
					Txn.CheckNum = sField
				End select
			Next
' e-Gold: sign of transaction depends on field 2
			If vFields(2) = "Payment Made" Then
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
			If PayeeLocation > 0 And Len(Txn.Payee) = 0 Then
				Txn.Payee = Trim(Mid(sMemo, PayeeLocation, PayeeLength))
			End If
			If Len(TxnDatePattern) > 0 Then
				vDateBits = ParseLineFixed(Txn.Memo, TxnDatePattern)
				If TypeName(vDateBits) = "Variant()" Then
					Txn.TxnDate = DateSerial(Year(Stmt.OpeningBalance.BalDate), CInt(vDateBits(2)), CInt(vDateBits(1))) _
						+ TimeSerial(CInt(vDateBits(3)), CInt(vDateBits(4)), 0)
					Txn.TxnDateValid = True
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

' eGold special: get statement dates. transactions are in ascending date order
			If Stmt.OpeningBalance.BalDate = NODATE Then
				Stmt.OpeningBalance.BalDate = Txn.BookDate
			End If
			Stmt.ClosingBalance.BalDate = Txn.BookDate

		End If
	Loop
	LoadTextFile = True
End Function
