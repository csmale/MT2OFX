' MT2OFX Input Processing Script For Societe Generale alternate CSV format

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/SocGen2-CSV.vbs 3     11/06/05 19:33 Colin $"

Const ScriptName = "SocGen2-CSV"
Const FormatName = "Societe Generale alternate CSV"
Const ParseErrorMessage = "Cannot parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

Const DebugRecognition = False	' enables debug code in recognition
Const BankCode = "SOGEFRPP"
Const CSVSeparator = ";"
Const NumFieldsExpected = 5
Const DateSequence = "DMY"	' must be DMY, MDY, or YMD
Const DateSeparator = "/"	' can be empty for dates in e.g. "yyyymmdd" format
Const CurrencyCode = ""	' default if not specified in file
Dim AccountNum 				' from properties
Const SkipHeaderLines = 2	' number of lines to skip before the transaction data
Const ColumnHeadersPresent = True	' are the column headers in the file?
Const DecimalSeparator = ","	' as used in amounts
Const MemoChunkLength = 0	' if memo field consists of fixed length chunks
Const TxnDatePattern = ".*(\d\d)/(\d\d)/(\d\d) (\d\d)H(\d\d).*"	' pattern to find transaction date in the memo
Const PayeeLocation = 1		' start of payee in memo
Const PayeeLength = 16		' length of payee in memo
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

' Declare fields in the order they appear in the file as an array of arrays. The inner arrays
' contain a field ID from the list above followed by the exact column header.
Dim aFields
aFields = Array( _
	Array(fldBookDate, "Date de l'opération"), _
	Array(fldPayee, "Libellé"), _
	Array(fldMemo, "Détail de l'écriture"), _
	Array(fldAmount, "Montant de l'opération"), _
	Array(fldCurrency, "Devise") _
)

' Dictionary to facilitate field lookup by field code
Dim FieldDict
Set FieldDict = CreateObject("Scripting.Dictionary")

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
'	LoadProperties ScriptName, aPropertyList
End Sub

' function DescriptiveName
' returns a string with a descriptive name of this script
Function DescriptiveName()
	DescriptiveName = FormatName
End Function

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
' special for SocGen2 - header line and some transactions have trailing delimiter
	If UBound(vFields) <> NumFieldsExpected And UBound(vFields) <> NumFieldsExpected+1 Then
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
	End If
	LogProgress ScriptName, "File Recognised"
	RecogniseTextFile = True
End Function

Function LoadTextFile()
	Dim sLine       ' holds a line
	Dim sPat        ' holds the match pattern
	Dim vFields     ' array of fields in the line
	Dim sType       ' record type
	Dim sAcct       ' last account number
	Dim Stmt        ' holds the current statement
	Dim sTmp		' temporary string
	Dim sTmp2       ' another one
	Dim vDateBits	' parts of date
	Dim iYear		' year
	Dim sOms1, sOms2, sOms3, sOms4, sOms5	' description lines
	Dim iSeq		' transaction sequence number
	Dim i
	Dim dBal		' temp balance date
	Dim aBal		' temp balance amount

	LoadTextFile = False
	sAcct = ""
' special for SocGen - balance is contained in first line
	sLine = ReadLine()
	vFields = ParseLineDelimited(sLine, CSVSeparator)
	If TypeName(vFields) <> "Variant()" Then
 MsgBox "not var array"
		Exit Function
	End If
	Set Stmt = NewStatement()
	Stmt.BranchName = left(vFields(1), 11)
	Stmt.Acct = mid(vFields(1), 12)
	Stmt.OpeningBalance.BalDate = NODATE
	Stmt.AvailableBalance.BalDate = NODATE
	Stmt.ClosingBalance.BalDate = ParseDate(vFields(5))				
	sTmp = vFields(6)
	Stmt.ClosingBalance.Ccy = Right(sTmp, 3)
	sTmp = Trim(Left(sTmp, Len(sTmp)-4))
	Stmt.ClosingBalance.Amt = ParseNumber(sTmp, DecimalSeparator)
	iSeq = 0
	Stmt.BankName = BankCode
	
	For i=2 To SkipHeaderLines
		sLine = ReadLine()
	Next
	If ColumnHeadersPresent Then
		sLine = ReadLine()
	End if
	Do While Not AtEOF()
		sLine = ReadLine()
		If Len(sLine) > 0 Then
			vFields = ParseLineDelimited(sLine, CSVSeparator)
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
' special for SocGen2 - header line and some transactions have trailing delimiter
			If UBound(vFields) <> NumFieldsExpected And UBound(vFields) <> NumFieldsExpected+1 Then
				Message True, True, "Wrong number of fields - " & UBound(vFields) & " - " & sLine, ScriptName
				Abort
				Exit function
			End If
	' set up new transaction, and start a new statement if the account # changes
			If FieldDict.Exists(fldAccountNum) Then
				If sAcct <> vFields(FieldDict(fldAccountNum)) Then
					Set Stmt = NewStatement()
		' this initialisation should be in the class constructor!! (fixed in 3.3.5)
					Stmt.OpeningBalance.BalDate = NODATE
					Stmt.AvailableBalance.BalDate = NODATE
					Stmt.ClosingBalance.BalDate = NODATE				
					iSeq = 0
					Stmt.BankName = BankCode
				End If
			Else
				If IsEmpty(Stmt) Then
					Set Stmt = NewStatement()
		' this initialisation should be in the class constructor!! (fixed in 3.3.5)
					Stmt.OpeningBalance.BalDate = NODATE
					Stmt.AvailableBalance.BalDate = NODATE
					Stmt.ClosingBalance.BalDate = NODATE				
					iSeq = 0
					Stmt.BankName = BankCode
				End If
			End If
			NewTransaction
			iSeq = iSeq + 1
			LastMemo = ""
			For i=1 To NumFieldsExpected
				Select Case aFields(i-1)(0)
				case fldSkip
				case fldAccountNum
					Stmt.Acct = vFields(i)
					sAcct = vFields(i)
				case fldCurrency
					Stmt.OpeningBalance.Ccy = vFields(i)
					Stmt.ClosingBalance.Ccy = vFields(i)
					Stmt.AvailableBalance.Ccy = vFields(i)
				case fldClosingBal
					Stmt.ClosingBalance.Amt = ParseNumber(vFields(i), DecimalSeparator)
				case fldAvailBal
					Stmt.AvailableBalance.Amt = ParseNumber(vFields(i), DecimalSeparator)
				case fldBookDate
					Txn.BookDate = ParseDate(vFields(i))
				case fldValueDate
					Txn.ValueDate = ParseDate(vFields(i))
				case fldAmtCredit
					Txn.Amt = Txn.Amt + Abs(ParseNumber(vFields(i), DecimalSeparator))
				case fldAmtDebit
					Txn.Amt = Txn.Amt - Abs(ParseNumber(vFields(i), DecimalSeparator))
				Case fldAmount
					Txn.Amt = ParseNumber(vFields(i), DecimalSeparator)
				case fldMemo
					Txn.Memo = vFields(i)
				Case fldBalanceDate
					dBal = ParseDate(vFields(i))
					If dBal > Stmt.ClosingBalance.BalDate Or Stmt.ClosingBalance.BalDate = NODATE Then
						Stmt.ClosingBalance.BalDate = dBal
						Stmt.AvailableBalance.BalDate = dBal
					End If
					If dBal < Stmt.OpeningBalance.BalDate Or Stmt.OpeningBalance.BalDate = NODATE Then
						Stmt.OpeningBalance.BalDate = dBal
					End If
				Case fldPayee
					Txn.Payee = vFields(i)
				End select
			Next
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
' special for SocGen2
			sMemo = Txn.Memo
			If StartsWith(sMemo, "FAC.CB.") Then
				Txn.TxnDate = ParseDate(Mid(Txn.Payee, 4,5) & "/" & "05")
				Txn.TxnDateValid = True
				Txn.Payee = Trim(Mid(sMemo, 29))
				Txn.TxnType = "POS"
			Elseif StartsWith(sMemo, "PRELEVEMENT ") Then
				Txn.Payee = Trim(Mid(sMemo, 25))
				Txn.TxnType = "DIRECTDEBIT"
			Elseif StartsWith(sMemo, "CHEQUE ") Then
				Txn.CheckNum = Trim(Mid(Txn.Payee, 8))
				Txn.Payee = ""
				Txn.TxnType = "CHECK"
			Elseif StartsWith(sMemo, "ECHEANCE ") Then
	' leave standard payee
			Elseif StartsWith(sMemo, "PRELEVT ") Then
	' leave standard payee
				Txn.TxnType = "DIRECTDEBIT"
			Elseif StartsWith(sMemo, "PRET ") Then
	' leave standard payee
			Elseif StartsWith(sMemo, "RET ECLAIR ") Then
				Txn.TxnType = "ATM"
				Txn.Payee = "Retrait"
	' transaction date/time will be found below
			Elseif StartsWith(sMemo, "TIP ") Then
	' leave standard payee
			Elseif StartsWith(sMemo, "VIREMENT ") Then
				Txn.Payee = Trim(Mid(sMemo, 24))
				Txn.Payee = Trim(Left(Txn.Payee, Len(Txn.Payee)-12))
				Txn.TxnType = "DIRECTDEP"
			Elseif StartsWith(sMemo, "VERSEMENT ") Then
	' leave standard payee
			Elseif StartsWith(sMemo, "VIR.PERM ") Then
	' leave standard payee
			End If

			If Len(TxnDatePattern) > 0 Then
				vDateBits = ParseLineFixed(Txn.Memo, TxnDatePattern)
				If TypeName(vDateBits) = "Variant()" Then
					Txn.TxnDate = DateSerial(CInt(vDateBits(3)), CInt(vDateBits(2)), CInt(vDateBits(1))) _
						+ TimeSerial(CInt(vDateBits(4)), CInt(vDateBits(5)), 0)
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

		End If
	Loop
' special for socgen
	Stmt.AvailableBalance.Ccy = ""

	LoadTextFile = True
End Function
