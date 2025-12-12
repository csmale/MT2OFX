' MT2OFX Input Processing Script for SNS Bank CSV format

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/SNSBank-Csv.vbs 7     18/12/06 21:37 Colin $"

Const ScriptName = "SNSBank-CSV"
Const FormatName = "SNS Bank CSV Formaat"
Dim NoOFXMessage : NoOFXMessage = "Van dit bestandstype kunt u geen OFC of OFX produceren omdat het geen saldoinformatie bevat." _
	& vbCrLf & vbCrLf & "Kies een ander uitvoerformaat zoals QIF."
Const NoOFXTitle = "Uitvoerformaat niet mogelijk."
Dim BadRecordTypeMessage : BadRecordTypeMessage = "Onbekend recordtype!!!" & vbCrLf & vbCrLf _
	& "Kan niet verder."
Dim BadRecordTypeTitle : BadRecordTypeTitle = ScriptName
Const ParseErrorMessage = "Kan regel niet ontleden."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

' fld#	formaat		inhoud
'	1	dd-mm-jjjj	journaaldatum
Const fldJournaalDatum = 1
'	2	9(10)		rekeningnummer
Const fldRekNr = 2
'	3	9(10)		tegenrekening
Const fldTgnRekNr = 3
'	4	X(28)		naam tegenrekening
Const fldNaam = 4
'	5	?			adres (toekomstig gebruik)
Const fldAdres = 5
'	6	9999XX		postcode (toekomstig gebruik)
Const fldPostcode = 6
'	7	X(24)		plaats
Const fldPlaats = 7
'	8	XXX			valutasoort rekening
Const fldRekValuta = 8
'	9	-9(9).99	saldo vóór deze transactie
Const fldVorigSaldo = 9
'	10	XXX			valutasoort transactie
Const fldTransValuta = 10
'	11	-9(9).99	transactiebedrag
Const fldBedrag = 11
'	12	dd-mm-jjjj	boekdatum
Const fldBoekDatum = 12
'	13	dd-mm-jjjj	valutadatum
Const fldValutaDatum = 13
'	14	9999		interne transactiecode
Const fldIntTransCode = 14
'	15	XXX			globale transactiecode
Const fldExtTransCode = 15
'	16	9(8)		transactievolgnummer
Const fldTransVolgNr = 16
'	17	X(16)		transactiekenmerk
Const fldKenmerk = 17
'	18	X(320)		omschrijving (max. 10 stukken van 32)
Const fldOmschrijving = 18
'   19	9			always seems to be "6"...???
Const fldSix = 19

Const BEADatePat = "(.*)(\d{1,2})\.(\d{1,2})\.(\d{4}) ([ \d]\d)u(\d{2}).*"

Dim LastMemo	' last non-blank memo field seen

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

Function ParseDate(sDate)
	Dim iYear, iMonth, iDay			' for dates, dd-mm-yyyy
	iYear = CInt(Mid(sDate,7,4))
	iMonth = CInt(Mid(sDate,4,2))
	iDay = CInt(Left(sDate,2))
	ParseDate = DateSerial(iYear, iMonth, iDay)
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
	If Len(Txn.FurtherInfo) > 0 Then
		Txn.FurtherInfo = Txn.FurtherInfo & Cfg.MemoDelimiter
	End If
	Txn.FurtherInfo = Txn.FurtherInfo & s
	LastMemo = s
End Sub

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
	Dim sLine
	RecogniseTextFile = False
	sLine = ReadLine()
	Dim vFields
	vFields = ParseLineDelimited(sLine, ",")
	If TypeName(vFields) <> "Variant()" Then
		Exit Function
	End If
	If UBound(vFields) <> 18 And Ubound(vFields) <> 19 Then
'		MsgBox "Num fields: " & CStr(UBound(vFields))
		Exit Function
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
	Dim vDateBits	' parts of date
	Dim iYear		' year

	LoadTextFile = False
	sAcct = ""
	Do While Not AtEOF()
		sLine = ReadLine()
		If Len(sLine) > 0 Then
			vFields = ParseLineDelimited(sLine, ",")
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
	' set up new transaction, and start a new statement if the account # changes
			If sAcct <> vFields(fldRekNr) Then
				Set Stmt = NewStatement()
				sAcct = vFields(fldRekNr)
				Stmt.Acct = Trim(sAcct)
				Stmt.BankName = "SNS Bank"
				Stmt.OpeningBalance.BalDate = ParseDate(vFields(fldJournaalDatum))
				Stmt.OpeningBalance.Ccy = vFields(fldRekValuta)
				Stmt.OpeningBalance.Amt = ParseNumber(vFields(fldVorigSaldo), ".")
				Stmt.ClosingBalance.Amt = Stmt.OpeningBalance.Amt
				Stmt.ClosingBalance.Ccy = vFields(fldRekValuta)
'				Stmt.ClosingBalance.Ccy = ""	' to force 00000000 as date in ledger bal
			End If
			NewTransaction
			LastMemo = ""
			Txn.Amt = ParseNumber(vFields(fldBedrag), ".")
			Stmt.ClosingBalance.Amt = Stmt.ClosingBalance.Amt + Txn.Amt
	' this will put the NAME of the payee into Txn.Payee.
			Txn.Payee = vFields(fldNaam)
	' the payee account number in vFields(5) is not useful!
			Txn.ValueDate = ParseDate(vFields(fldValutaDatum))
			Txn.BookDate = ParseDate(vFields(fldJournaalDatum))
			Stmt.ClosingBalance.BalDate = Txn.BookDate
			Txn.IsReversal = False
			Txn.TxnType = TransType(vFields(fldExtTransCode))
	' not happy about TxnDate - the documentation is a bit ambiguous and we don't
	' really have enough test data to validate it properly
			If vFields(fldExtTransCode) = "BEA" Then
				vDateBits = ParseLineFixed(vFields(fldOmschrijving), BEADatePat)
				If TypeName(vDateBits) = "Variant()" Then
					Txn.TxnDate = DateSerial(vDateBits(4), CInt(vDateBits(3)), CInt(vDateBits(2))) _
						+ TimeSerial(CInt(vDateBits(6)), CInt(vDateBits(6)), 0)
					Txn.TxnDateValid = True
					Txn.Payee = vDateBits(1)
				End If
			End If
' cash withdrawals - special string for payee
			If vFields(fldExtTransCode) = "GEA" Then
				Txn.Payee = "Kasopname"
			End If
' the description is all in one field already
			ConcatMemo Trim(vFields(fldOmschrijving))

' sort out a transaction ID - date, sequence are provided!
			Txn.FITID = CStr(Year(Txn.BookDate)) & _
				Right("0" & CStr(Month(Txn.BookDate)), 2) & _
				Right("0" & CStr(Day(Txn.BookDate)), 2) & _
				"." & vFields(fldTransVolgNr)
		End If
	Loop
	LoadTextFile = True
End Function

Function TransType(sCode)
	Select Case sCode
	Case "ACC"	'	Acceptgirobetaling
		TransType = "OTHER"
	Case "AF"	'	Afboeking
		TransType = "OTHER"
	Case "AFB"	'	Afbetalen
		TransType = "OTHER"
	Case "BEA"	'	Betaalautomaat
		TransType = "OTHER"
	Case "BIJ"	'	Bijboeking
		TransType = "OTHER"
	Case "BTL"	'	Buitenlandse Overboeking
		TransType = "OTHER"
	Case "CHP"	'	Chipknip
		TransType = "OTHER"
	Case "CHQ"	'	Cheque
		TransType = "OTHER"
	Case "COR"	'	Correctie
		TransType = "OTHER"
	Case "DIV"	'	Diversen
		TransType = "OTHER"
	Case "EFF"	'	Effectenboeking
		TransType = "OTHER"
	Case "ETC"	'	Euro traveller cheques
		TransType = "OTHER"
	Case "GBK"	'	GiroBetaalkaart
		TransType = "OTHER"
	Case "GEA"	'	Geldautomaat
		TransType = "OTHER"
	Case "INC"	'	Incasso
		TransType = "OTHER"
	Case "IOB"	'	Interne Overboeking
		TransType = "OTHER"
	Case "KAS"	'	Kas post
		TransType = "OTHER"
	Case "KNT"	'	Kosten/provisies
		TransType = "OTHER"
	Case "KST"	'	Kosten/provisies
		TransType = "OTHER"
	Case "OVB"	'	Overboeking
		TransType = "OTHER"
	Case "PRM"	'	Premies
		TransType = "OTHER"
	Case "PRV"	'	Provisies
		TransType = "OTHER"
	Case "RNT"	'	Rente
		TransType = "OTHER"
	Case "STO"	'	Storno
		TransType = "OTHER"
	Case "TEL"	'	Telefonische Overboeking
		TransType = "OTHER"
	Case "VV"	'	Vreemde valuta
		TransType = "OTHER"
	Case Else
		TransType = "OTHER"
	End Select
End Function
