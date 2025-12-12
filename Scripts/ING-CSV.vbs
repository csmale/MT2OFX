' MT2OFX Input Processing Script for ING Bank NL CSV format

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/ING-CSV.vbs 2     17/07/04 0:19 Colin $"

Const ScriptName = "INGBankNL-CSV"
Const FormatName = "ING Bank (NL) Comma-Separated Formaat"
Const ParseErrorMessage = "Kan regel niet ontleden."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

' fld#	inhoud
'	1	rekeningnummer
Const fldRekNr = 1
'	2	boekdatum
Const fldBoekDatum = 2
'	3	bij/af
Const fldBijAf = 3
'	4	bedrag
Const fldBedrag = 4
'	5	tegenrekening
Const fldTgnRek = 5
'	6	omschrijving
Const fldOmschrijving = 6

Sub Initialise()
    LogProgress ScriptName, "Initialise"
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

' ING uses dd-mm-yyyy in fldBoekDatum,
'  dd.mm.yyyy hh:mm in txn info for POS/ATM
Function ParseDate(sDate)
	Dim iYear, iMonth, iDay			' for dates
	Dim iHour, iMin, iSec
	iDay = CInt(Left(sDate,2))
	iMonth = CInt(Mid(sDate,4,2))
	iYear = CInt(Mid(sDate,7,4))
	ParseDate = DateSerial(iYear, iMonth, iDay)
	If Len(sDate) > 10 Then
		iHour = CInt(Mid(sDate, 12, 2))
		iMin = CInt(Mid(sDate, 15, 2))
		iSec = 0
		ParseDate = ParseDate + TimeSerial(iHour, iMin, iSec)
	End If
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

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
	Dim vFields
	Dim sLine
	RecogniseTextFile = False
	sLine = ReadLine()
	vFields = ParseLineDelimited(sLine, ";")
	If TypeName(vFields) <> "Variant()" Then
		Exit Function
	End If
	If UBound(vFields) <> 6 Then
		Exit Function
	End If
	If vFields(fldBijAf) <> "Bij" And vFields(fldBijAf) <> "Af" Then
		Exit Function
	End If
	LogProgress ScriptName, "File Recognised - " & FormatName
	RecogniseTextFile = True
End Function

Function LoadTextFile()
	Dim sLine       ' holds a line
	Dim vFields     ' array of fields in the line
	Dim sAcct       ' last account number
	Dim Stmt        ' holds the current statement
	Dim sTmp		' temporary string
	Dim dBal		' temp date for check num
	Dim iSeq		' txn sequence num
	Dim dSeq		' date for sequence
	Dim iPos		' temp for tidying up payee

	LoadTextFile = False
	sAcct = ""
	Do While Not AtEOF()
		sLine = ReadLine()
		sLine = Replace(sLine, Chr(10), "")
		If Len(sLine) > 0 Then
			vFields = ParseLineDelimited(sLine, ";")
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			If UBound(vFields) <> 6 Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
	MsgBox Asc(left(sLine,1)) & " " & Len(sLine)
				Abort
				Exit Function
			End If
			
	' set up new transaction, and start a new statement if the account # changes
			If sAcct <> vFields(fldRekNr) Then
				Set Stmt = NewStatement()
				sAcct = vFields(fldRekNr)
				Stmt.Acct = Trim(sAcct)
				Stmt.BankName = "INGBankNL"
				Stmt.OpeningBalance.Ccy = "EUR"	' no info in file
				Stmt.OpeningBalance.BalDate = ParseDate(vFields(fldBoekDatum))
				Stmt.ClosingBalance.Ccy = ""	' to force 00000000 as date in ledger bal
			End If
			NewTransaction
			Txn.Amt = ParseNumber(vFields(fldBedrag), ",")
			If vFields(fldBijAf) = "Af" Then
				Txn.Amt = -Txn.Amt
			End If
			Txn.ValueDate = ParseDate(vFields(fldBoekDatum))
			Txn.BookDate = ParseDate(vFields(fldBoekDatum))
			Stmt.ClosingBalance.BalDate = Txn.ValueDate	' file goes backwards!
			Txn.IsReversal = False
			Txn.FurtherInfo = Trim(vFields(fldOmschrijving))

' default transaction type
			If Txn.Amt > 0 Then
				Txn.TxnType = "DEP"
			Else
				Txn.TxnType = "PAYMENT"
			End If
			
' get transaction date out of description
			If StartsWith(Txn.FurtherInfo, "BETAALAUTOMAAT ") Then
				Txn.TxnDate = ParseDate(Mid(Txn.FurtherInfo, 16, 16))
				Txn.TxnDateValid = True
				Txn.Payee = Mid(Txn.FurtherInfo, 36)	
				Txn.TxnType = "POS"
			ElseIf StartsWith(Txn.FurtherInfo, "GELDAUTOMAAT ") Then
				Txn.TxnDate = ParseDate(Mid(Txn.FurtherInfo, 14, 16))
				Txn.TxnDateValid = True
				Txn.Payee = "Kasopname"
				Txn.TxnType = "ATM"
			End If
			
' special strings for other payees
			If StartsWith(Txn.FurtherInfo, "CHIPKNIP ") Then
				Txn.Payee = "Chipknip"
				Txn.TxnType = "ATM"
			Elseif StartsWith(Txn.FurtherInfo, "DEBETRENTE VAN ") Then
				Txn.Payee = "Rente"
				Txn.TxnType = "INT"
			Elseif StartsWith(Txn.FurtherInfo, "CREDITRENTE VAN ") Then
				Txn.Payee = "Rente"
				Txn.TxnType = "INT"
			Elseif StartsWith(Txn.FurtherInfo, "BIJDRAGE EUROPAS ") Then
				Txn.Payee = "Bank"
				Txn.TxnType = "SRVCHG"
			End If
			
' sometimes we don't get a payee! If we use the tegenrekening, Money
'	can sort it out
			If Txn.Payee = "" Then
				If vFields(fldTgnRek) = "-" Then
					Txn.Payee = "Onbekend"
				Else
					Txn.Payee = vFields(fldTgnRek)
				End If
			End If

			iPos = InStr(Txn.Payee, " ,PAS ")
			If iPos <> 0 Then
				Txn.Payee = Left(Txn.Payee, iPos-1)
			End If
			iPos = InStr(Txn.Payee, " RUNNUMMER_BGC: ")
			If iPos <> 0 Then
				Txn.Payee = Left(Txn.Payee, iPos-1)
			End If
			Txn.Payee = Trim(Left(Txn.Payee, 32))		' OFX: max len 32

' sort out a transaction ID
			dBal = Txn.BookDate
			If dSeq = dBal Then
				iSeq = iSeq + 1
			Else
				iSeq = 1
				dSeq = dBal
			End If
			Txn.CheckNum = CStr(DatePart("yyyy", dSeq)) _
				& Right("00" & CStr(DatePart("y", dSeq)), 3) _
				& "." & CStr(iSeq)
		End If
	Loop
	LoadTextFile = True
End Function

