' MT2OFX Input Processing Script for Postbank NL Asc/CSV formats

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/Postbank-AscCSV.vbs 11    30/01/11 13:32 Colin $"

Const ScriptName = "Postbank-AscCSV"
Const FormatName = "Postbank (NL) Asc/CSV (Girotel Online/Offline) Formaat"
Const ParseErrorMessage = "Kan regel niet ontleden."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

' fld#	posn	len	inhoud
'	1	1-10	10	rekeningnummer
Const fldRekNr = 1
'	2	11-18	8	datum yyyymmdd
Const fldDatum = 2
'	3	19-21	3	txn code XXX
Const fldTransCode = 3
'	4	22-24	3	txn volgnummer binnen dag nnn
Const fldTransNr = 4
'	5	25-34	10	tegenrekening (onbetrouwbaar)
Const fldTgnRekNr = 5
'	6	35-66	32	omschrijving 1 = naam
Const fldNaam = 6
'	7	67-86	20	omschrijving 2 = aftrekpost
Const fldAftrekPost = 7
'	8	87-98	12	bedrag (centen)
Const fldBedrag = 8
'	9	99-100	2	Af/Bij + "M" ?
' Const fldAfBij = 9
Dim fldAfBij
'	10	101-132	32	omschrijving 3
' Const fldOmschrijving1 = 10
Dim fldOmschrijving1
'	11	133-164	32	omschrijving 4
' Const fldOmschrijving2 = 11
Dim fldOmschrijving2
'	12	165-196	32	omschrijving 5
' Const fldOmschrijving3 = 12
Dim fldOmschrijving3
'	13	197-228	32	omschrijving 6
' Const fldOmschrijving4 = 13
Dim fldOmschrijving4
'	14	229-260	32	omschrijving 7
' Const fldOmschrijving5 = 14
Dim fldOmschrijving5
'	15	261-263	3	valuta (EUR)
' Const fldValuta = 15
Dim fldValuta

' pattern for Girotel Online ASC
Const MainPat = "(\d{10})(\d{8})(.{3})(\d{3})(\d{10})(.{32})(.{20})(\d{12})(.).(.{32})(.{32})(.{32})(.{32})(.{32})(...)"
' pattern for Girotel Offline ASC
Const OfflinePat = "(\d{10})(\d{8})(.{3})(\d{3})(\d{10})(.{32})(.{20})(\d{12})(...)(.).(.{32})?(.{32})?(.{32})?(.{32})?(.{32})?"
' this holds the pattern to be used
Dim LinePat
' this indicates whether we are processing a CSV file
Dim bIsCSV

Const BADatePat = "PASNR .{7} (\d{2})-(\d{2})-(\d{2}) (\d{2}) UUR (\d{2}).*"
' PASNR 999X999 dd-mm-yy hh UUR mm
Const GMDatePat = " (\d{2})-(\d{2})-(\d{2}) (\d{2}):(\d{2}).*"
' dd-mm-yy 16:52 999X999  9999999 

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
	Dim iYear, iMonth, iDay			' for dates
	iYear = CInt(Left(sDate,4))
	iMonth = CInt(Mid(sDate,5,2))
	iDay = CInt(Mid(sDate,7,2))
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
	Dim vFields
	Dim sLine
	RecogniseTextFile = False
	bIsCSV = False
	sLine = ReadLine()
' check for CSV format
	vFields = ParseLineDelimited(sLine, ",")
	If TypeName(vFields) = "Variant()" Then
		If UBound(vFields) >= 12 Then
		If UBound(vFields) = 12 And vFields(12) = "EUR" Then	' Girotel Online
			fldAfBij = 9
			fldOmschrijving1 = 11
			fldOmschrijving2 = -1
			fldOmschrijving3 = -1
			fldOmschrijving4 = -1
			fldOmschrijving5 = -1
			fldValuta = 12
			bIsCSV = True
			RecogniseTextFile = True
			LogProgress ScriptName, "File Recognised - CSV Girotel Online"
			Exit Function
		Elseif UBound(vFields) > 12 And vFields(9) = "EUR" Then ' Girotel Offline
' ********* aarghhh! variable number of fields. needs fixing!
			fldValuta = 9
			fldAfBij = 10
			fldOmschrijving1 = 12
			fldOmschrijving2 = 13
			fldOmschrijving3 = 14
			fldOmschrijving4 = 15
			fldOmschrijving5 = 16
			bIsCSV = True
			RecogniseTextFile = True
			LogProgress ScriptName, "File Recognised - CSV Girotel Offline"
			Exit Function
		End If
		End If
	End If
	
	If Len(sLine) = 263 And (Right(sLine, 3) = "EUR" Or Right(sLine, 3) = "   ") Then	' Girotel Online
		LinePat = MainPat
		fldAfBij = 9
		fldOmschrijving1 = 10
		fldOmschrijving2 = 11
		fldOmschrijving3 = 12
		fldOmschrijving4 = 13
		fldOmschrijving5 = 14
		fldValuta = 15
	Else	' assume Girotel Offline
		LinePat = OfflinePat
		fldAfBij = 10
		fldOmschrijving1 = 11
		fldOmschrijving2 = 12
		fldOmschrijving3 = 13
		fldOmschrijving4 = 14
		fldOmschrijving5 = 15
		fldValuta = 9
	End If
' test parse the line - if it fails, we don't recognise the file
	vFields = ParseLineFixed(sLine, LinePat)
	If TypeName(vFields) <> "Variant()" Then
'		MsgBox "Unexpected failure to parse line"
		Exit Function
	End If
    If UBound(vFields) < 0 Then
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
	Dim sTmp2       ' another one
	Dim sTmp3       ' and again!
	Dim vDateBits	' parts of date
	Dim iYear		' year
	Dim sOms1, sOms2, sOms3, sOms4, sOms5	' description lines

	LoadTextFile = False
	sAcct = ""
	Do While Not AtEOF()
		sLine = ReadLine()
		If Len(sLine) > 0 Then
			If bIsCSV Then
				vFields = ParseLineDelimited(sLine, ",")
			Else
				vFields = ParseLineFixed(sLine, LinePat)
			End If
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
            If UBound(vFields) < 12 Then
                Exit Function
            End If
			vFields(fldTransCode) = Trim(vFields(fldTransCode))
	' set up new transaction, and start a new statement if the account # changes
			If sAcct <> vFields(fldRekNr) Then
				Set Stmt = NewStatement()
				sAcct = vFields(fldRekNr)
				Stmt.Acct = Trim(sAcct)
				Stmt.BankName = "Postbank"
' 20060915 CS: currency now hard coded if it is not provided in the file (it was always EUR anyway)			
				If Len(Trim(vFields(fldValuta))) = 0 Then
					Stmt.OpeningBalance.Ccy = "EUR"
				Else
					Stmt.OpeningBalance.Ccy = Trim(vFields(fldValuta))
				End If
				Stmt.OpeningBalance.BalDate = ParseDate(vFields(fldDatum))
				Stmt.ClosingBalance.Ccy = ""	' to force 00000000 as date in ledger bal
			End If
			NewTransaction
			LastMemo = ""
			If bIsCSV Then
				Txn.Amt = ParseNumber(vFields(fldBedrag), ".")
			Else
				Txn.Amt = CDbl(vFields(fldBedrag)) / 100.0
			End If
			If vFields(fldAfBij) <> "B" Then	' could be "A" (Af) or blank - also DR
				Txn.Amt = -Txn.Amt
			End If
	' this will put the NAME of the payee into Txn.Payee.
			Txn.Payee = Trim(Replace(vFields(fldNaam),">"," "))
			If vFields(fldTransCode) = "BA" Then
	' strip off the 8 digits plus space (what are they for?)
				Txn.Payee = Trim(Mid(Txn.Payee, 9))
	' capture trimmed full field in memo
				ConcatMemo Txn.Payee
	' strip off trailing digits (often branch number)
	' comment out this line if required!
		'		Txn.Payee = TrimTrailingDigits(Txn.Payee)
			Else
	' capture full field in memo
				ConcatMemo Trim(vFields(fldNaam))
			End If
	' the payee account number in vFields(5) is not useful!
			Txn.ValueDate = ParseDate(vFields(fldDatum))
			Txn.BookDate = ParseDate(vFields(fldDatum))
			Stmt.ClosingBalance.BalDate = Txn.ValueDate
			Txn.IsReversal = False
			Txn.TxnType = TransType(vFields(fldTransCode), Txn.Amt)
			If vFields(fldTransCode) = "BA" Then
				vDateBits = ParseLineFixed(vFields(fldOmschrijving1), BADatePat)
			ElseIf vFields(fldTransCode) = "GM" Then
				vDateBits = ParseLineFixed(vFields(fldOmschrijving1), GMDatePat)
			Else
				vDateBits = False	' as long as it is not an array!
			End If
' the following code works for all cases as long as the date/time fields are in the right order
			If TypeName(vDateBits) = "Variant()" Then
            If UBound(vDateBits) >= 5 Then
                Txn.TxnDate = DateSerial(vDateBits(3), CInt(vDateBits(2)), CInt(vDateBits(1))) _
                    + TimeSerial(CInt(vDateBits(4)), CInt(vDateBits(5)), 0)
                Txn.TxnDateValid = True
            End If
			End If
' cash withdrawals - special string for payee
			If vFields(fldTransCode) = "GM" Or vFields(fldTransCode) = "PK" Then
				Txn.Payee = "Kasopname"
			End If

			If bIsCSV And Ubound(vFields) <= 12 Then
				If fldOmschrijving1 <= UBound(vFields) Then
					ConcatMemo Trim(Mid(vFields(fldOmschrijving1),1,32))
					ConcatMemo Trim(Mid(vFields(fldOmschrijving1),33,32))
					ConcatMemo Trim(Mid(vFields(fldOmschrijving1),65,32))
					ConcatMemo Trim(Mid(vFields(fldOmschrijving1),97,32))
					ConcatMemo Trim(Mid(vFields(fldOmschrijving1),129,32))
				End If
			Else
				If fldOmschrijving1 <= UBound(vFields) Then
					ConcatMemo Trim(vFields(fldOmschrijving1))
				End If
				If fldOmschrijving2 <= UBound(vFields) Then
					ConcatMemo Trim(vFields(fldOmschrijving2))
				End If
				If fldOmschrijving3 <= UBound(vFields) Then
					ConcatMemo Trim(vFields(fldOmschrijving3))
				End If
				If fldOmschrijving4 <= UBound(vFields) Then
					ConcatMemo Trim(vFields(fldOmschrijving4))
				End If
				If fldOmschrijving5 <= UBound(vFields) Then
					ConcatMemo Trim(vFields(fldOmschrijving5))
				End If
			End If
' very special checks!
			If bIsCSV And Ubound(vFields) <= 12 Then
				If fldOmschrijving1 <= UBound(vFields) Then
					sTmp = Trim(Left(vFields(fldOmschrijving1),32))
					sTmp2 = Trim(Mid(vFields(fldOmschrijving1), 33, 32))
					sTmp3 = Trim(Mid(vFields(fldOmschrijving1), 65, 32))
				Else
					sTmp = ""
					sTmp2 = ""
					sTmp3 = ""
				End If
			Else
				sTmp = Trim(vFields(fldOmschrijving1))
				sTmp2 = Trim(vFields(fldOmschrijving2))
				sTmp3 = Trim(vFields(fldOmschrijving3))
			End If

' klantspecifiek
'			If StartsWith(sTmp, "UWV ") _
'			Or StartsWith(sTmp, "SPORT-EN") _
'			Or StartsWith(sTmp, "BANKGIROLOTERIJ") Then

			If StartsWith(sTmp, "UWV ") _
			Or StartsWith(sTmp, "BANKGIROLOTERIJ") Then
				Txn.Payee = sTmp
			ElseIf StartsWith(sTmp2, "STAATSLOTERIJ") Then
				Txn.Payee = "Staatsloterij"
			Else
			' direct debits almost always have the payee name in the last memo field
			' CS 20041124: ... unless its the new Interpay layout...
				If vFields(fldTransCode) = "IC" Then
					If fldOmschrijving3 <= UBound(vFields) Then
						If sTmp = "INTERPAY BEANET BV" And sTmp2 = "Uw pinbetaling bij" Then
							Txn.Payee = sTmp3
						Else
							Txn.Payee = LastMemo
						End If
					Else
						Txn.Payee = LastMemo
					End If
            ElseIf StartsWith(Txn.Memo, "KN: ") Then
'msgbox sTmp & vbcrlf & stmp2 & vbcrlf & stmp3
                Txn.Payee = sTmp
				End If
			End If
' klantspecifiek
'			If Txn.Payee = "GIROTEL/GIRONET ABONNEMENTSGELD" Then
'				Txn.Payee = "POSTBANK N.V."
'			End If

' sort out a transaction ID - date, sequence are provided!
			Txn.CheckNum = CStr(Year(Txn.BookDate)) & _
				Right("0" & CStr(Month(Txn.BookDate)), 2) & _
				Right("0" & CStr(Day(Txn.BookDate)), 2) & _
				"." & vFields(fldTransNr)
		End If
	Loop
	LoadTextFile = True
End Function

Function TransType(PostbankCode, Amt)
	Select Case PostbankCode
	Case "AC"
		TransType = "DIRECTDEBIT"
	Case "BA"
		TransType = "POS"
	Case "CH"
		TransType = "CHECK"
	Case "GB"
		TransType = "CHECK"
	Case "GM"
		TransType = "ATM"
	Case "IC"
		TransType = "DIRECTDEBIT"
	Case "PK"
		TransType = "CASH"
	Case "ST"
		TransType = "DEP"
	Case "VZ"
		TransType = "DIRECTDEP"
	Case Else
		TransType = "OTHER"
	End Select
End Function
'	other known codes:
'	RES, TAN, GIN, GEW, DV, EUR, FL, GF, GT, OV, PO, TA
