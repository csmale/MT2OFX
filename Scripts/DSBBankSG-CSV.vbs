' MT2OFX Input Processing Script DSB Bank (Singapore) CSV format

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/DSBBankSG-CSV.vbs 7     27/06/10 21:45 Colin $"

Const ScriptName = "DSBBankSG-CSV"
Const FormatName = "DSB Bank (Singapore) CSV Format"
Const ParseErrorMessage = "Cannot parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

Const DebugRecognition = False	' enables debug code in recognition
Const BankCode = "DBSSSGSG"
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
Const CSVSeparator = ","
Const TxnLinePattern = ""
Const NumFieldsExpected = 7
Const DateSequence = "DMY"	' must be DMY, MDY, or YMD
Const DateSeparator = " "	' can be empty for dates in e.g. "yyyymmdd" format
Const InvertSign = False	' make credits into debits etc
Const CurrencyCode = "SGD"	' default if not specified in file
Dim AccountNum		' default if not specified in file
Const SkipHeaderLines = 5	' number of lines to skip before the transaction data
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
Const fldFITID = 16

' Declare fields in the order they appear in the file as an array of arrays. The inner arrays
' contain a field ID from the list above followed by the exact column header.
' Transaction Date,Reference,Debit Amount,Credit Amount,Transaction Ref1,Transaction Ref2,Transaction Ref3
' Transaction Date,Value Date,Statement Code,Reference,Debit Amount,Credit Amount,Client Reference,Additional Reference
Dim aFields
aFields = Array( _
	Array(fldTransactionDate, "Transaction Date"), _
	Array(fldSkip, "Reference"), _
	Array(fldAmtDebit, "Debit Amount"), _
	Array(fldAmtCredit, "Credit Amount"), _
	Array(fldMemo, "Transaction Ref1"), _
	Array(fldMemo, "Transaction Ref2"), _
	Array(fldMemo, "Transaction Ref3"), _
	Array(fldMemo, ""), _
	Array(fldSkip, ""), _
	Array(fldSkip, "") _
)
' Transaction Date,Value Date,Statement Code,Reference,Debit Amount,Credit Amount,
'  Client Reference,Additional Reference
Dim aFields2
aFields2 = Array( _
	Array(fldTransactionDate, "Transaction Date"), _
	Array(fldValueDate, "Value Date"), _
	Array(fldSkip, "Statement Code"), _
	Array(fldSkip, "Reference"), _
	Array(fldAmtDebit, "Debit Amount"), _
	Array(fldAmtCredit, "Credit Amount"), _
	Array(fldMemo, "Client Reference"), _
	Array(fldMemo, "Additional Reference"), _
	Array(fldSkip, ""), _
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
	Array("AXACompte", "Numéro compte", _
		"Le numéro de compte pour AXA Belgique en format 000-000000-00.", _
		ptString) _
	)

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
			MsgBox "not var array"
		End If
		Exit Function
	End If
' 20090310 CS: new variant
	If UBound(vFields) >=2 Then
		If vFields(2) = "Value Date" Then aFields=aFields2
	End If
' 20060130 CS: DSB Bank has trailing fields
	If UBound(vFields) < NumFieldsExpected Then
		If DebugRecognition Then
			MsgBox "wrong number of fields - got " & UBound(vFields) & ", expected " _
			& NumFieldsExpected & " - " & sLine
		End If
		Exit function
	End If
	If ColumnHeadersPresent Then
		For i=1 To UBound(vFields)
			If vFields(i) <> aFields(i-1)(1) Then
				If DebugRecognition Then
					MsgBox "field " & CStr(i) & " " & aFields(i-1)(1) & " instead of " & vFields(i)
				End If
				Exit function
			End if
		Next
	Else
' pattern-match the first row
		For i=1 To UBound(vFields)
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
						sPat = "[+-]?[ 0-9,]*(\.[0-9]*)?"
					Else
						sPat = "[+-]?[ 0-9\.]*(,[0-9]*)?"
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
				If DebugRecognition Then
					MsgBox "Field " & i & " (" & sField & ") failed to match"
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
	Dim iTmp		' temporary integer (duh)
	Dim vDateBits	' parts of date
	Dim iSeq		' transaction sequence number
	Dim i
	Dim dBal		' temp balance date
	Dim sField		' field value being processed
	Dim dMaxDate	' latest txn/book date - if we don't have a statement date

	LoadTextFile = False
	sAcct = ""
	Set Stmt = NewStatement()
	Stmt.OpeningBalance.BalDate = NODATE
	Stmt.OpeningBalance.Ccy = CurrencyCode
	Stmt.AvailableBalance.BalDate = NODATE
	Stmt.AvailableBalance.Ccy = CurrencyCode
	Stmt.ClosingBalance.BalDate = NODATE				
	Stmt.ClosingBalance.Ccy = CurrencyCode
	Stmt.BankName = BankCode
	For i=1 To SkipHeaderLines
		sLine = ReadLine()
		If StartsWith(sLine, "Account Details For:") Then
			AccountNum = Trim(Mid(sLine, InStrRev(sLine, " ")+1))
		Elseif StartsWith(sLine, "Statement as at:") Then
			Stmt.ClosingBalance.BalDate = ParseDate(Right(sLine, 11))
			Stmt.AvailableBalance.BalDate = Stmt.ClosingBalance.BalDate
		Elseif StartsWith(sLine, "Available Balance:") Then
			Stmt.AvailableBalance.Amt = ParseNumber(Mid(sLine, InStr(sLine, ",")+1),".")
		Elseif StartsWith(sLine, "Ledger Balance:") Then
			Stmt.ClosingBalance.Amt = ParseNumber(Mid(sLine, InStr(sLine, ",")+1),".")
		End If
	Next
	Stmt.Acct = AccountNum
	
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
' 20060130 CS: DSB Bank has trailing fields
			If UBound(vFields) < NumFieldsExpected Then
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
					dMaxDate = NODATE
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
					dMaxDate = NODATE
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
					If Txn.BookDate <> NODATE Then
						If dMaxDate = NODATE Or Txn.BookDate > dMaxDate Then
							dMaxDate = Txn.BookDate
						End If
					End If
				case fldValueDate
					Txn.ValueDate = ParseDate(sField)
				Case fldTransactionDate
					Txn.TxnDate = ParseDate(sField)
					Txn.TxnDateValid = (Txn.TxnDate <> NODATE)
					If Txn.TxnDate <> NODATE Then
						If dMaxDate = NODATE Or Txn.TxnDate > dMaxDate Then
							dMaxDate = Txn.TxnDate
						End If
					End If
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
' DSB: set book date to txn Date
			Txn.BookDate = Txn.TxnDate
			If Txn.BookDate <> NODATE Then
				If dMaxDate = NODATE Or Txn.BookDate > dMaxDate Then
					dMaxDate = Txn.BookDate
				End If
			End If
' DSB: set up FITID based on book Date
			If dBal <> Txn.BookDate Then
				iSeq = 1
				dBal = Txn.BookDate
			End If
			Txn.FITID = CStr(Year(Txn.BookDate)) _
                & "." & Right("000" & DatePart("y", Txn.BookDate), 3) _
                & "." & Right("000" & Cstr(iSeq), 3)
' DSB: track starting Date
			If Stmt.OpeningBalance.BalDate = NODATE Then
				Stmt.OpeningBalance.BalDate = Txn.BookDate
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
                If UBound(vDateBits) >= 4 Then
                    Txn.TxnDate = DateSerial(Year(Stmt.OpeningBalance.BalDate), CInt(vDateBits(2)), CInt(vDateBits(1))) _
                        + TimeSerial(CInt(vDateBits(3)), CInt(vDateBits(4)), 0)
                    Txn.TxnDateValid = True
                End If
				End If
			End If
' DSB: type-dependent payee processing
' 20090310 CS: new variant transaction type
			If aFields(2)(0) = fldSkip Then
				Select Case vFields(3)
				Case "POS"
					Txn.TxnType = "POS"
					Txn.Payee = Txn.Memo
				Case "CHQ", "CSH"
					If Txn.Amt < 0 And IsNumeric(vFields(4)) Then
						Txn.TxnType = "CHECK"
						Txn.CheckNum = vFields(4)
					End If
				Case "ATM"
					If Txn.Amt < 0 Then
						Txn.TxnType = "ATM"
					End If
				End Select
			Else
				Select Case Trim(vFields(2))
				Case "POS","ITR"
					Txn.Payee = Trim(vFields(5))
				Case "AWL", "CAM", "ASC"
					Txn.Payee = Trim(vFields(6))
				Case "SI","IBG","SAL","ITX"
					sTmp = Trim(vFields(5))
					iTmp = InStr(sTmp, ":")
					If iTmp > 0 Then
						sTmp = Mid(sTmp, iTmp+1)
					End If
					Txn.Payee = Trim(sTmp)
				Case "BILL"
					iTmp = InStr(Txn.Memo, " ")
					If iTmp > 0 Then
						Txn.Payee = Trim(Left(Txn.Memo, iTmp-1))
					End If
            Case "MST"
                Txn.Payee = Trim(Left(vFields(5), Len(vFields(5))-11))
                iTmp = InStr(Txn.Payee, "*")
                If iTmp > 0 Then
                    Txn.Payee = Trim(Mid(Txn.Payee, iTmp+1))
                End If
				End Select
' DSB: transaction type
				Select Case Trim(vFields(2))
				Case "ASC": Txn.TxnType = "SRVCHG"
				Case "ADP","AWL": Txn.TxnType = "ATM"
				End Select
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

' OFX Transaction Types
' Type        Description
' CREDIT      Generic credit
' DEBIT       Generic debit
' INT         Interest earned or paid
'             Note: Depends on signage of amount
' DIV         Dividend
' FEE         FI fee
' SRVCHG      Service charge
' DEP         deposit
' ATM         ATM debit or credit
'             Note: Depends on signage of amount
' POS         Point of sale debit or credit
'             Note: Depends on signage of amount
' XFER        Transfer
' CHECK       Check
' PAYMENT     Electronic payment
' CASH        Cash withdrawal
' DIRECTDEP   Direct deposit
' DIRECTDEBIT Merchant initiated debit
' REPEATPMT   Repeating payment/standing order
' OTHER       Other
' =======================================
