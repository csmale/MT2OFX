' MT2OFX Input Processing Script Smile Bank (UK) web clipboard format

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/SmileUK-WEB.vbs 5     14/10/09 22:33 Colin $"

Const ScriptName = "SmileUK-WEB"
Const FormatName = "Smile Bank (UK) Web Clipboard Format"
Const ParseErrorMessage = "Cannot parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

Const DebugRecognition = False	' enables debug code in recognition
Const BankCode = "CPBKGB11"
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
Dim CSVSeparator: CSVSeparator = vbTab
Const TxnLinePattern = ""
Const NumFieldsExpected = 5
Const DateSequence = "DMY"	' must be DMY, MDY, or YMD
Const DateSeparator = "/"	' can be empty for dates in e.g. "yyyymmdd" format
Const InvertSign = False	' make credits into debits etc
Const CurrencyCode = "GBP"	' default if not specified in file
Dim AccountNum		         ' default if not specified in file
Dim BranchID
Const SkipHeaderLines = 26	' number of lines to skip before the transaction data
Const ColumnHeadersPresent = True	' are the column headers in the file?
Const DecimalSeparator = "."	' as used in amounts
Const MemoChunkLength = 0	' if memo field consists of fixed length chunks
Const TxnDatePattern = ".*(\d\d):(\d\d)([A-Z]{3})(\d\d).*"	' pattern to find transaction date in the memo
' e.g. "16:49DEC21" - notice NO YEAR
' Const TxnDatePattern = ""
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
' date 	transaction 	money in 	money out 	balance
Dim aFields
aFields = Array( _
	Array(fldBookDate, "date", "date ?"), _
	Array(fldMemo, "transaction", "transaction ?"), _
	Array(fldAmtCredit, "money in", "money in ?"), _
	Array(fldAmtDebit, "money out", "money out ?"), _
	Array(fldClosingBal, "balance", "balance ?") _
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
' aPropertyList = Array( _
'	Array("AXACompte", "Numéro compte", _
'		"Le numéro de compte pour AXA Belgique en format 000-000000-00.", _
'		ptString) _
'	)
aPropertyList = Array()

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
        If Left(sLine, 4) = "date" Then
            exit For
        End If
	Next
	If AtEOF() Then
		Exit Function
	End If
    If Left(sLine, 4) <> "date" Then
        sLine = ReadLine()
    End If
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
	If UBound(vFields) <> NumFieldsExpected Then
		If DebugRecognition Then
			MsgBox "wrong number of fields - got " & UBound(vFields) & ", expected " _
			& NumFieldsExpected & " - " & sLine
		End If
		Exit function
	End If
	If ColumnHeadersPresent Then
		For i=1 To NumFieldsExpected
            If UBound(aFields(i-1)) > 1 Then
                If Not StringMatches(vFields(i), aFields(i-1)(2)) Then
                    If DebugRecognition Then
                        MsgBox "field " & CStr(i) & ", got '" & vFields(i) & "', expected pattern '" & aFields(i-1)(2) & "'"
                    End If
                    Exit Function
                End If
            Else
                If vFields(i) <> aFields(i-1)(1) Then
                    If DebugRecognition Then
                        MsgBox "field " & CStr(i) & ", got '" & vFields(i) & "', expected '" & aFields(i-1)(2) & "'"
                    End If
                    Exit Function
                End if
            End If
		Next
	Else
' pattern-match the first row
		For i=1 To NumFieldsExpected
			sField = Trim(vFields(i))
			If UBound(aFields(i-1)) > 1 Then
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
	Dim vDateBits	' parts of date
	Dim iSeq		' transaction sequence number
	Dim i
	Dim dBal		' temp balance date
	Dim sField		' field value being processed
	Dim dMaxDate	' latest txn/book date - if we don't have a statement Date

	LoadTextFile = False
	sAcct = ""
	For i=1 To SkipHeaderLines
		sLine = ReadLine()
' smile: capture account number
		If StartsWith(sLine, "account number") Then
			AccountNum = Trim(Replace(Mid(sLine, 17), vbTab, ""))
		End If
      If StartsWith(sLine, "sort code") Then
        BranchID = Trim(Replace(Mid(sLine, 10), vbTab, ""))
      End If
        If StartsWith(sLine, "date") Then
            Exit For
        End If
	Next
'	If StartsWith(sLine, "date") Then
'        sLine = ReadLine()
'	End If
	Do While Not AtEOF()
		sLine = Trim(ReadLine())
		If Len(sLine) = 0 Then Exit Do
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
			If UBound(vFields) <> NumFieldsExpected Then
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
               Stmt.BranchName = BranchID
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
               Stmt.BranchName = BranchID
					Stmt.Acct = AccountNum
					dMaxDate = NODATE
				End If
			End If
			NewTransaction
' need GUIDs as days can get split across pages
'			Txn.FITID = MakeGUID()
' don't need this any more
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
' smile: amounts include pound sign
					Stmt.ClosingBalance.Amt = ParseNumber(Replace(sField,"£",""), DecimalSeparator)
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
					Txn.Amt = Txn.Amt + Abs(ParseNumber(Replace(sField,"£",""), DecimalSeparator))
				case fldAmtDebit
					Txn.Amt = Txn.Amt - Abs(ParseNumber(Replace(sField,"£",""), DecimalSeparator))
				Case fldAmount
					Txn.Amt = ParseNumber(sField, DecimalSeparator)
				Case fldChequeNum
					Txn.CheckNum = sField
				case fldMemo
					ConcatMemo Trim(sField)
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
				End Select
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
' smile txn types
			If Txn.Memo = "BROUGHT FORWARD" Then
				Stmt.OpeningBalance.Amt = Txn.Amt
				Txn.Amt = 0
				Txn.TxnType = "OTHER"
			End If
			If StartsWith(Txn.Memo, "LINK ") Then
				Txn.TxnType = "ATM"
			Elseif StartsWith(Txn.Memo, "INTEREST ") Then
				Txn.TxnType = "INT"
			Elseif IsNumeric(Txn.Memo) Then
				Txn.TxnType = "CHECK"
				Txn.CheckNum = Txn.Memo
			End If
						
			Dim sMemo
' find the payee, transaction type and txn date if we can
			sMemo = Txn.Memo
			If Txn.TxnType = "PAYMENT" And Len(Txn.Payee) = 0 Then
				Txn.Payee = Trim(TrimTrailingDigits(sMemo))
			End If
			If Len(TxnDatePattern) > 0 Then
				vDateBits = ParseLineFixed(Txn.Memo, TxnDatePattern)
' smile: parts are hour, minute, month (3 letters), Day
' no year!
				If UBound(vDateBits) > 0 Then
					sTmp = vDateBits(4) & "-" & vDateBits(3) & "-" & CStr(Year(Txn.BookDate))
					Txn.TxnDate = ParseDateEx(sTmp, "DMY", "-")
					If Txn.TxnDate <> NODATE Then
						Txn.TxnDate = NearestYear(Txn.TxnDate, Txn.BookDate)
						Txn.TxnDate = Txn.TxnDate + TimeSerial(CInt(vDateBits(1)), CInt(vDateBits(2)), 0)
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
			Stmt.ClosingBalance.BalDate = Txn.BookDate
			If Stmt.OpeningBalance.BalDate = NODATE Then
				Stmt.OpeningBalance.BalDate = Txn.BookDate
			End If
		End If
' no available balance in statement
		Stmt.AvailableBalance.Ccy = ""
	Loop
	LoadTextFile = True
End Function

' sort out a year for dDate (which hasn't got a reliable year part)
' if the date is in january and the base is in december, use next year
' if the date is in december and the base is in january, use last year
Function NearestYear(dDate, dBase)
	Dim iYear
	iYear = Year(dBase)
	If Month(dDate) = 1 Then
		If Month(dBase) = 12 Then
			iYear = iYear + 1
		End If
	Elseif Month(dDate) = 12 Then
		If Month(dBase) = 1 Then
			iYear = iYear - 1
		End If
	End If
	NearestYear = DateSerial(iYear, Month(dDate), Day(dDate)) + TimeSerial(Hour(dDate), Minute(dDate), Second(dDate))
End Function
