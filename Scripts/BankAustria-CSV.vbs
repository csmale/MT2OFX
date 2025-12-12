' MT2OFX Input Processing Script Bank Austria Creditanstalt AG CSV format
Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/BankAustria-CSV.vbs 8     11/01/10 19:23 Colin $"
Const ScriptName = "BankAustria-CSV"
Const FormatName = "Bank Austria Creditanstalt AG"
Const ParseErrorMessage = "Cannot parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

Const BankCode = "BKAUATWW"
Const CSVSeparator = ";"
Const NumFieldsExpected = 13
Const DateSequence = "DMY"	' must be DMY, MDY, or YMD
Const DateSeparator = "."	' can be empty for dates in e.g. "yyyymmdd" format
Const SkipHeaderLines = 0	' number of lines to skip before the transaction data
Const ColumnHeadersPresent = True	' are the column headers in the file?
Dim DecimalSeparator: DecimalSeparator = ","	' as used in amounts
Const MemoChunkLength = 58	' if memo field consists of fixed length chunks
Const TxnDatePattern = ".*(\d\d)\.(\d\d)\. ?UM (\d\d)\.(\d\d).*"	' pattern to find transaction date in the memo

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

' Declare fields in the order they appear in the file
Dim aFields
Dim aFieldsG
' German version:
' Downloaddatum;Kontonummer;Kontowortlaut;Kontowaehrung;Valutasaldo;Disposaldo;
' Kontostand;Buchungsdatum;Valutadatum;Buchungswaehrung;Eingang;Ausgang;Buchungstext
aFieldsG = Array( _
	Array(fldBalanceDate, "Downloaddatum"), _
	Array(fldAccountNum, "Kontonummer"), _
	Array(fldSkip, "Kontowortlaut"), _
	Array(fldCurrency, "Kontowaehrung"), _
	Array(fldClosingBal, "Valutasaldo"), _
	Array(fldAvailBal, "Disposaldo"), _
	Array(fldClosingBal, "Kontostand"), _
	Array(fldBookDate, "Buchungsdatum"), _
	Array(fldValueDate, "Valutadatum"), _
	Array(fldSkip, "Buchungswaehrung"), _
	Array(fldAmtCredit, "Eingang"), _
	Array(fldAmtDebit, "Ausgang"), _
	Array(fldMemo, "Buchungstext") _	
)
' English version:
' Downloaddate;Account number;Account name;Account currency;Cleared balance;Drawing limit;_
' Account balance;Booking date;Cleared date;Currency;In;Out;Account entry text
Dim aFieldsE
aFieldsE = Array( _
	Array(fldBalanceDate, "Downloaddate"), _
	Array(fldAccountNum, "Account number"), _
	Array(fldSkip, "Account name"), _
	Array(fldCurrency, "Account currency"), _
	Array(fldClosingBal, "Cleared balance"), _
	Array(fldAvailBal, "Drawing limit"), _
	Array(fldClosingBal, "Account balance"), _
	Array(fldBookDate, "Booking date"), _
	Array(fldValueDate, "Cleared date"), _
	Array(fldSkip, "Currency"), _
	Array(fldAmtCredit, "In"), _
	Array(fldAmtDebit, "Out"), _
	Array(fldMemo, "Account entry text") _	
)
' have we been through the recognition process?
Dim bRecognitionDone: bRecognitionDone = False

' Dictionary to facilitate field lookup by field code
Dim FieldDict
Set FieldDict = CreateObject("Scripting.Dictionary")

'20050415;'xxxx2484400;Name of accountholder;EUR;1873,72;5973,72;
'	1873,72;13.04.2005;12.04.2005;EUR;;-15,52;KIK          0063  K1 12.04.UM 16.04

Dim LastMemo	' last non-blank memo field seen

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

Function StartsWith(s, Prefix)
	StartsWith = (Left(s,Len(Prefix)) = Prefix)
End Function

Function ParseDate(sDate)
	ParseDate = ParseDateEx(sDate, DateSequence, DateSeparator)
End Function

Function TrimTrailingDigits(s)
	Dim r
	Set r=New regexp
	r.Global = False
	r.Pattern = "^(.*?) *\d+$"
	Dim m
	Set m=r.Execute(s)
	If m.Count = 0 Then
		TrimTrailingDigits = s
	Else
		TrimTrailingDigits = m(0).SubMatches(0)
	End If
End Function

Sub ConcatMemo(s)
	If s = "" Then
		Exit Sub
	End If
	If Len(Txn.Memo) > 0 Then
		Txn.Memo = Txn.Memo & Cfg.MemoDelimiter
	End If
	Txn.Memo = Txn.Memo & s
	LastMemo = s
End Sub

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
	Dim vFields
	Dim sLine
	Dim i
	RecogniseTextFile = False
	For i=1 To SkipHeaderLines
		sLine = ReadLine()
	Next
	sLine = ReadLine()
	vFields = ParseLineDelimited(sLine, CSVSeparator)
	If TypeName(vFields) <> "Variant()" Then
' MsgBox "not var array"
		Exit Function
	End If
	If UBound(vFields) <> NumFieldsExpected Then
' MsgBox "wrong number of fields - " & CStr(UBound(vFields)+1) & " - " & sline
		Exit function
	End If
' switch language if needed
	If vFields(1) = "Downloaddate" Then
		aFields = aFieldsE
		DecimalSeparator = "."
	Else
		aFields = aFieldsG
		DecimalSeparator = ","
	End If

' fill field lookup dictionary
' NB: only the last occurrence is remembered!
	For i=0 To UBound(aFields)
		FieldDict(aFields(i)(0)) = i+1
	Next

	If ColumnHeadersPresent Then
		For i=1 To NumFieldsExpected
			If vFields(i) <> aFields(i-1)(1) Then
' MsgBox "field " & CStr(i) & " " & aFields(i-1)(1) & " instead of " & vFields(i)
				Exit function
			End if
		Next
	End If
	LogProgress ScriptName, "File Recognised"
	bRecognitionDone = True
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

	LoadTextFile = False

' if we haven't been through recognition, just check it now	
	If Not bRecognitionDone Then
		If Not RecogniseTextFile() Then
			Abort
			Exit Function
		End If
		Rewind
	End If
	
	sAcct = ""
	For i=1 To SkipHeaderLines
		sLine = ReadLine()
	Next
	If ColumnHeadersPresent Then
		sLine = ReadLine()
	End if
	Do While Not AtEOF()
		sLine = ReadLine()
		If Len(sLine) > 0 Then
' special for Bank Austria Creditanstalt
			sLine = Replace(sLine, ";'", ";")
			vFields = ParseLineDelimited(sLine, CSVSeparator)
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
' txn can have more fields than expected - these days the memo can be split over multiple fields.
			If UBound(vFields) < NumFieldsExpected Then
				Message True, True, "Wrong number of fields - " & CStr(UBound(vFields)+1) & " - " & sLine, ScriptName
				Abort
				Exit function
			End If
	' set up new transaction, and start a new statement if the account # changes
			If sAcct <> vFields(FieldDict(fldAccountNum)) Then
				Set Stmt = NewStatement()
	' this initialisation should be in the class constructor!! (fixed in 3.3.5)
				Stmt.OpeningBalance.BalDate = NODATE
				Stmt.AvailableBalance.BalDate = NODATE
				Stmt.ClosingBalance.BalDate = NODATE				
				iSeq = 0
				Stmt.BankName = BankCode
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
' special for BankAustriaCreditanstalt - this date field has a different format to the others!
'					dBal = ParseDate(vFields(i))
					dBal = ParseDateEx(vFields(i), "YMD", "") - 1
					If dBal > Stmt.ClosingBalance.BalDate Or Stmt.ClosingBalance.BalDate = NODATE Then
						Stmt.ClosingBalance.BalDate = dBal
						Stmt.AvailableBalance.BalDate = dBal
					End If
					If dBal < Stmt.OpeningBalance.BalDate Or Stmt.OpeningBalance.BalDate = NODATE Then
						Stmt.OpeningBalance.BalDate = dBal
					End If
				End select
			Next
' see if there are any extra memo fields
			For i=NumFieldsExpected+1 To UBound(vFields)
				ConcatMemo vFields(i)
			Next
			
			If Txn.Amt < 0 Then
				Txn.TxnType = "PAYMENT"
			Else
				Txn.TxnType = "DEP"
			End If
			
			Dim sMemo
' find the payee, transaction type and txn date if we can
			sMemo = Trim(Txn.Memo)
' special for Bank Austria Creditanstalt
			If StartsWith(sMemo, "Gutschrift ") Then
				Txn.Payee = Trim(Mid(sMemo, 14, 40))
				Txn.TxnType = "DIRECTDEP"
			Elseif StartsWith(sMemo, "Lastschrift ") Then
				Txn.Payee = Trim(Mid(sMemo, 15, 40))
				Txn.TxnType = "DIRECTDEBIT"
			Elseif StartsWith(sMemo, "EZE-Lastschrift ") Then
				Txn.Payee = Trim(Mid(sMemo, 19, 40))
				Txn.TxnType = "DIRECTDEBIT"
			Elseif StartsWith(sMemo, "BANKOMAT ") Then
				Txn.Payee = "Bankomat"
				Txn.TxnType = "ATM"
			Elseif StartsWith(sMemo, "ABHEBUNG AM AUTOMAT ") Then
				Txn.Payee = "Bankomat"
				Txn.TxnType = "ATM"
			Elseif StartsWith(sMemo, "Jahrespreis ") Then
				Txn.TxnType = "SRVCHG"
			Elseif StartsWith(sMemo, "Kontoführung") Then
				Txn.TxnType = "SRVCHG"
			Elseif StartsWith(sMemo, "Spesen ") Then
				Txn.TxnType = "SRVCHG"
			Elseif Instr(sMemo, "%") > 5 And Instr(sMemo, "%") < 15 Then
				Txn.TxnType = "INT"
			Elseif StartsWith(sMemo, "Online-Auftrag vom ") Then
				i = InStr(sMemo, "Empfänger:")
				If i>0 Then
					Txn.Payee = Trim(Mid(sMemo, i+11, 50))
				End If
			Else
				Txn.Payee = Trim(Left(sMemo, 13))
			End If
			If Len(TxnDatePattern) > 0 Then
				vDateBits = ParseLineFixed(sMemo, TxnDatePattern)
				If TypeName(vDateBits) = "Variant()" Then
                If UBound(vDateBits) >= 4 Then
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

		End If
	Loop
	LoadTextFile = True
End Function
