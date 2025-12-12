' MT2OFX Input Processing Script Dexia Bank Belgium CSV format

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/DexiaBE-CSV.vbs 15    8/07/08 22:34 Colin $"

Const ScriptName = "DexiaBE-CSV"
Const FormatName = "Dexia Belgium CSV"
Const ParseErrorMessage = "Cannot parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

Const DebugRecognition = False
Const BankCode = "GKCCBEBB"
Const CSVSeparator = ";"
Const NumFieldsExpected = 11
Const NumFieldsExpectedMax = 15
Const DateSequence = "DMY"	' must be DMY, MDY, or YMD
Const DateSeparator = "/"	' can be empty for dates in e.g. "yyyymmdd" format
Const SkipHeaderLines = 14	' number of lines to skip before the transaction data
Const ColumnHeadersPresent = True	' are the column headers in the file?
Const DecimalSeparator = ","	' as used in amounts
Const MemoChunkLength = 0	' if memo field consists of fixed length chunks
Const TxnDatePattern = ""	' pattern to find transaction date in the memo
Const PayeeLocation = 0		' start of payee in memo
Const PayeeLength = 0		' length of payee in memo

Dim PayeeField, PayeeField2	' field NUMBERS where the payee can be found. if PayeeField is empty Then
							' PayeeField2 is tried. Zero field numbers are skipped.
PayeeField = 0: PayeeField2 = 0

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

'Compte;Date de comptabilisation;Numéro sur extrait de compte;
'Compte contrepartie;Nom et adresse contrepartie;Rue et numéro;Code postal et localité;
'Communication;Date valeur;Montant;Code devise

' Declare fields in the order they appear in the file as an array of arrays. The inner arrays
' contain a field ID from the list above followed by the exact column header.
Dim aFields
Dim aFieldsFr
aFieldsFr = Array( _
	Array(fldAccountNum, "Compte"), _
	Array(fldBookDate, "Date de comptabilisation"), _
	Array(fldSkip, "Numéro sur extrait de compte"), _
	Array(fldSkip, "Compte contrepartie"), _
	Array(fldMemo, "Nom et adresse contrepartie"), _
	Array(fldMemo, "Rue et numéro"), _
	Array(fldMemo, "Code postal et localité"), _
	Array(fldMemo, "Communication"), _
	Array(fldValueDate, "Date valeur"), _
	Array(fldAmount, "Montant"), _
	Array(fldCurrency, "Code devise"), _
	Array(fldSkip, "") _	
)

' Old version:
'Rekening;Boekingsdatum;Nummer van de verrichting op het rekeninguittreksel;Rekening tegenpartij;Naam en adres tegenpartij;
'Straat en nummer;Postcode en gemeente;Mededeling;Valutadatum;Bedrag;Muntcode;
Dim aFieldsNl
aFieldsNl = Array( _
	Array(fldAccountNum, "Rekening"), _
	Array(fldBookDate, "Boekingsdatum"), _
	Array(fldSkip, "Nummer van de verrichting op het rekeninguittreksel"), _
	Array(fldSkip, "Rekening tegenpartij"), _
	Array(fldMemo, "Naam en adres tegenpartij"), _
	Array(fldMemo, "Straat en nummer"), _
	Array(fldMemo, "Postcode en gemeente"), _
	Array(fldMemo, "Mededeling"), _
	Array(fldValueDate, "Valutadatum"), _
	Array(fldAmount, "Bedrag"), _
	Array(fldCurrency, "Muntcode"), _
	Array(fldSkip, "") _	
)
' New version:
'Rekening;Boekingsdatum;Nummer afschrift;Rekening tegenpartij;Naam tegenpartij;
'Straat en nummer;Postcode en gemeente;Transactie;Valutadatum;Saldo;Muntcode
'Rekening;Boekingsdatum;Nummer afschrift;Rekening tegenpartij;Naam tegenpartij;
'Straat en nummer;Postcode en gemeente;Transactie;Valutadatum;Bedrag;Muntcode
' NB: the "Saldo" field is still really the transaction amount...
Dim aFieldsNl2
aFieldsNl2 = Array( _
	Array(fldAccountNum, "Rekening"), _
	Array(fldBookDate, "Boekingsdatum"), _
	Array(fldSkip, "Nummer afschrift"), _
	Array(fldSkip, "Rekening tegenpartij"), _
	Array(fldMemo, "Naam tegenpartij"), _
	Array(fldMemo, "Straat en nummer"), _
	Array(fldMemo, "Postcode en gemeente"), _
	Array(fldMemo, "Transactie"), _
	Array(fldValueDate, "Valutadatum"), _
	Array(fldAmount, "Saldo", "Saldo|Bedrag"), _
	Array(fldCurrency, "Muntcode"), _
	Array(fldSkip, "") _	
)

'Another new version:
'Rekening;Boekingsdatum;Rekening tegenpartij;Naam en adres tegenpartij;Straat en nummer;Postcode en gemeente;
'Mededeling;Valutadatum;Bedrag;Muntcode;
Dim aFieldsNl3
aFieldsNl3 = Array( _
	Array(fldSkip, "Rekening"), _
	Array(fldBookDate, "Boekingsdatum"), _
	Array(fldSkip, "Rekening tegenpartij"), _
	Array(fldMemo, "Naam en adres tegenpartij"), _
	Array(fldMemo, "Straat en nummer"), _
	Array(fldMemo, "Postcode en gemeente"), _
	Array(fldMemo, "Mededeling"), _
	Array(fldValueDate, "Valutadatum"), _
	Array(fldAmount, "Bedrag"), _
	Array(fldCurrency, "Muntcode"), _
	Array(fldSkip, "") _	
)

' Alternative yet another new version:
' Rekening;Boekingsdatum;Nummer afschrift;Transactienummer;Rekening tegenpartij;Naam tegenpartij;
' Straat en nummer;Postcode en localiteit;Transactie;Valutadatum;Bedrag;Muntcode;BIC;Landcode
Dim aFieldsNl4
aFieldsNl4 = Array( _
	Array(fldAccountNum, "Rekening"), _
	Array(fldBookDate, "Boekingsdatum"), _
	Array(fldSkip, "Nummer afschrift"), _
	Array(fldSkip, "Transactienummer"), _
	Array(fldSkip, "Rekening tegenpartij"), _
	Array(fldMemo, "Naam tegenpartij"), _
	Array(fldMemo, "Straat en nummer"), _
	Array(fldMemo, "Postcode en localiteit"), _
	Array(fldMemo, "Transactie"), _
	Array(fldValueDate, "Valutadatum"), _
	Array(fldAmount, "Bedrag"), _
	Array(fldCurrency, "Muntcode"), _
	Array(fldSkip, "BIC"), _
	Array(fldSkip, "Landcode"), _
	Array(fldSkip, "") _	
)


' Assume French to start with
aFields = aFieldsFr

' Dictionary to facilitate field lookup by field code
Dim FieldDict
Set FieldDict = CreateObject("Scripting.Dictionary")

Sub Initialise()
    LogProgress ScriptVersion, "Initialise"
	If Not CheckVersion() Then
		Abort
	End If
' Initialise dictionary of month names
	InitialiseMonths MonthNames
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
		If StartsWith(sLine, "Rekening;") Or StartsWith(sLine, "Compte;") Then
			Exit For
		End If
	Next
	If AtEOF() Then
		Exit Function
	End If
'	sLine = ReadLine()
	vFields = ParseLineDelimited(sLine, CSVSeparator)
	If TypeName(vFields) <> "Variant()" Then
		If DebugRecognition Then
			MsgBox "not var array",,ScriptName
		End If
		Exit Function
	End If
	If UBound(vFields) < NumFieldsExpected Or UBound(vFields) > NumFieldsExpectedMax Then
		If DebugRecognition Then
			MsgBox "wrong number of fields - got " & UBound(vFields) & ", expected " _
			& NumFieldsExpected & "-" & NumFieldsExpectedMax & " - " & sLine,,ScriptName
		End If
		Exit Function
	End If
	If ColumnHeadersPresent Then
' DexiaBE: switch to Dutch...
		If vFields(1) = "Rekening" Then
			If vFields(4) = "Transactienummer" Then
				aFields = aFieldsNl4
			ElseIf vFields(3) = "Rekening tegenpartij" Then
				aFields = aFieldsNl3
			ElseIf vFields(3) = "Nummer afschrift" Then
				aFields = aFieldsNl2
			Else
				aFields = aFieldsNl
			End If
			If DebugRecognition Then
				MsgBox "Switching to Dutch",,ScriptName
			End If
		Else
			aFields = aFieldsFr
		End If
		For i=1 To UBound(vFields)
			If UBound(aFields(i-1)) > 1 Then
				If Not StringMatches(vFields(i), aFields(i-1)(2)) Then
					If DebugRecognition Then
						MsgBox "Field " & CStr(i) & ": '" & vFields(i) & "' does not match '" & aFields(i-1)(2) & "'",,ScriptName
					End If
					Exit Function
				End If
			Else
				If vFields(i) <> aFields(i-1)(1) Then
					If DebugRecognition Then
						MsgBox "Field " & CStr(i) & ": got '" & vFields(i) & ", expecting '" & aFields(i-1)(1) & "'",,ScriptName
					End If
					Exit function
				End If
			End If
		Next
	End If
' fill field lookup dictionary
' NB: only the last occurrence is remembered!
	For i=0 To UBound(aFields)
		FieldDict(aFields(i)(0)) = i+1
	Next
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
	Dim sField		' field being processed

	LoadTextFile = False
	sAcct = ""
	dBal = NODATE
	For i=1 To SkipHeaderLines
		sLine = ReadLine()
' special for dexia
		If Len(sLine) > 0 And Left(sLine,1) <> CSVSeparator Then
			vFields = ParseLineDelimited(sLine, CSVSeparator)
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			If vFields(1) = "Solde" Or vFields(1) = "Saldo" Or vFields(1) = "Laatste saldo" Then
				sTmp = vFields(2)
				sTmp = Left(sTmp, InStr(sTmp, " ")-1)	' trim off " EUR"
	'			MsgBox "Balance: " & sTmp
				aBal = ParseNumber(sTmp, ",")
	'			MsgBox "Balance: " & aBal
			Elseif StartsWith(sLine, "Date/heure du solde;") Or StartsWith(sLine, "Datum/uur saldo;") Or StartsWith(sLine, "Datum/uur van het laatste saldo;") Then
				sTmp = vFields(2)
	'			MsgBox "Balance date: " & sTmp
				dBal = ParseDate(sTmp)
	'			MsgBox "Balance date: " & dBal
			ElseIf StartsWith(sLine, "Rekening;") Or StartsWith(sLine, "Compte;") Then
				Exit For
			End If
		End If
	Next

'Rekening;Boekingsdatum;Rekening tegenpartij;
	If ColumnHeadersPresent Then
		PayeeField = 5
		PayeeField2 = 4
'		sLine = ReadLine()
		If StartsWith(sLine, "Rekening;") Then
			If vFields(4) = "Transactienummer" Then
				aFields = aFieldsNl4
				PayeeField = 6
				PayeeField2 = 5
			ElseIf StartsWith(sLine, "Rekening;Boekingsdatum;Rekening tegenpartij;") Then
				aFields = aFieldsNl3
			Elseif StartsWith(sLine, "Rekening;Boekingsdatum;Nummer afschrift;")  Then
				aFields = aFieldsNl2
			Else
				aFields = aFieldsNl
			End If
			If DebugRecognition Then
				MsgBox "Switching to Dutch",,ScriptName
			End If
		Else
			aFields = aFieldsFr
		End If
	End If
	
	Do While Not AtEOF()
		sLine = ReadLine()
		If Len(sLine) > 0 And Left(sLine,1) <> CSVSeparator Then
' Dexia: place names can legitimately start with a single quote and these are not wrapped in double quotes
			sLine = FixQuotes(sLine)
'			MsgBox sLine
			vFields = ParseLineDelimited(sLine, CSVSeparator)
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
' special for dexia - some lines have an extra empty field
			If UBound(vFields) < NumFieldsExpected Or UBound(vFields) > (NumFieldsExpectedMax) Then
				Message True, True, "Wrong number of fields - " & CStr(UBound(vFields)+1) & " - " & sLine, ScriptName
				Abort
				Exit Function
			End If
	' set up new transaction, and start a new statement if the account # changes
			If FieldDict.Exists(fldAccountNum) Then
				If DebugRecognition Then MsgBox "acct num field exists"
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
				If DebugRecognition Then MsgBox "no acct num fld"
				If IsEmpty(Stmt) Then
					If DebugRecognition Then MsgBox "stmt is empty"
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
			If Len(TxnDatePattern) > 0 Then
				vDateBits = ParseLineFixed(Txn.Memo, TxnDatePattern)
				If TypeName(vDateBits) = "Variant()" Then
					Txn.TxnDate = DateSerial(Year(Stmt.OpeningBalance.BalDate), CInt(vDateBits(2)), CInt(vDateBits(1))) _
						+ TimeSerial(CInt(vDateBits(3)), CInt(vDateBits(4)), 0)
					Txn.TxnDateValid = True
				End If
			End If
' special for dexia - payee in col 5. If missing try for account number in col 4.
			Txn.Payee = Trim(vFields(PayeeField))
			If Txn.Payee = "" Then
				Txn.Payee = Trim(vFields(PayeeField2))
			End If

' tidy up the memo
			sMemo = Txn.Memo
			If StartsWith(sMemo, "OPVRAGING ") Then
				Txn.TxnType = "ATM"
			ElseIf InStr(sMemo, " DOMICILIERING ") <> 0 Then
				Txn.TxnType = "DIRECTDEBIT"
			ElseIf InStr(sMemo, "AANKOOP MISTER CASH ") <> 0 Then
				Txn.TxnType = "POS"
			End If			
			If MemoChunkLength > 0 Then
				Txn.Memo = ""
				For i=1 To Len(sMemo) Step MemoChunkLength
					ConcatMemo Trim(Mid(sMemo, i, MemoChunkLength))
				Next
			End If

		End If
	Loop
' special for dexia - set up statement here due to balance being in the header lines
	Stmt.AvailableBalance.Ccy = ""
	Stmt.ClosingBalance.BalDate = dBal				
	Stmt.ClosingBalance.Amt = aBal
	LoadTextFile = True
End Function

' Dexia: place names can legitimately start with a single quote and these are not wrapped in double quotes

Const FixQuotesPattern = "(.*;)('[^;]*)(;.*)"
Const FixQuotesReplace = "$1""$2""$3"

Function FixQuotes(sLine)
	Dim r
	Set r=New RegExp
	r.Global = True
	r.Pattern = FixQuotesPattern
	r.IgnoreCase = True
	FixQuotes = r.Replace(sLine, FixQuotesReplace)
End Function
