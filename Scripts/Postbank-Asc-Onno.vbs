' MT2OFX Input Processing Script for Postbank NL Asc format

Option Explicit
Const ScriptName = "Postbank-Asc"
Const FormatName = "Postbank (NL) Asc Formaat"
Dim NoOFXMessage : NoOFXMessage = "Van dit bestandstype kunt u geen OFC of OFX produceren omdat het geen saldoinformatie bevat." _
	& vbCrLf & vbCrLf & "Kies een ander uitvoerformaat zoals QIF."
Const NoOFXTitle = "Uitvoerformaat niet mogelijk."
Dim BadRecordTypeMessage : BadRecordTypeMessage = "Onbekend recordtype!!!" & vbCrLf & vbCrLf _
	& "Kan niet verder."
Dim BadRecordTypeTitle : BadRecordTypeTitle = ScriptName
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
Const fldAfTrekPost = 7
'	8	87-98	12	bedrag (centen)
Const fldBedrag = 8
'	9	99-100	2	Af/Bij + "M" ?
Const fldAfBij = 9
'	10	101-132	32	omschrijving 3
Const fldOmschrijving1 = 10
'	11	133-164	32	omschrijving 4
Const fldOmschrijving2 = 11
'	12	165-196	32	omschrijving 5
Const fldOmschrijving3 = 12
'	13	197-228	32	omschrijving 6
Const fldOmschrijving4 = 13
'	14	229-260	32	omschrijving 7
Const fldOmschrijving5 = 14
'	15	261-263	3	valuta (EUR)
Const fldValuta = 15

Const MainPat = "(\d{10})(\d{8})(.{3})(\d{3})(\d{10})(.{32})(.{20})(\d{12})(..)(.{32})(.{32})(.{32})(.{32})(.{32})(...)"
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
	Dim sLine
	RecogniseTextFile = False
	sLine = ReadLine()
	If Len(sLine) <> 263 Then
'		MsgBox "Line len: " & CStr(Len(sLine))
		Exit Function
	End If
	If Trim(Mid(sLine, 67, 20)) <> "" Then
'		MsgBox "stuff not found in first line"
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
	If Session.OutputFileType = "OFX" Or Session.OutputFileType = "OFC" Then
		MsgBox NoOFXMessage, vbOkOnly + vbCritical, NoOFXTitle
		Abort
		Exit Function
	End If
	sAcct = ""
	Do While Not AtEOF()
		sLine = ReadLine()
		If Len(sLine) > 0 Then
			vFields = ParseLineFixed(sLine, MainPat)
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
				Stmt.OpeningBalance.Ccy = vFields(fldValuta)
				Stmt.ClosingBalance.Ccy = vFields(fldValuta)
			End If
			NewTransaction
			LastMemo = ""
			Txn.Amt = CDbl(vFields(fldBedrag)) / 100.0
			If Left(vFields(fldAfBij),1) = "A" Then
				Txn.Amt = -Txn.Amt
			End If
	' this will put the NAME of the payee into Txn.Payee.
			Txn.Payee = Trim(Replace(vFields(fldNaam),">"," "))
			If vFields(fldTransCode) = "BA " Then
	' strip off the 8 digits plus space (what are they for?)
				Txn.Payee = Trim(Mid(Txn.Payee, 9))
	' capture trimmed full field in memo
				ConcatMemo Txn.Payee
	' strip off trailing digits (often branch number)
	' comment out this line if required!
				Txn.Payee = TrimTrailingDigits(Txn.Payee)
			Else
	' capture full field in memo
				ConcatMemo Trim(vFields(fldNaam))
			End If
	' the payee account number in vFields(5) is not useful!
			Txn.ValueDate = ParseDate(vFields(fldDatum))
			Txn.BookDate = ParseDate(vFields(fldDatum))
			Stmt.ClosingBalance.BalDate = Txn.ValueDate
			Txn.IsReversal = False
			If vFields(fldTransCode) = "BA " Then
				vDateBits = ParseLineFixed(vFields(fldOmschrijving1), BADatePat)
			ElseIf vFields(fldTransCode) = "GM " Then
				vDateBits = ParseLineFixed(vFields(fldOmschrijving1), GMDatePat)
			Else
				vDateBits = False	' as long as it is not an array!
			End If
' the following code works for all cases as long as the date/time fields are in the right order
			If TypeName(vDateBits) = "Variant()" Then
				Txn.TxnDate = DateSerial(vDateBits(3), CInt(vDateBits(2)), CInt(vDateBits(1))) _
					+ TimeSerial(CInt(vDateBits(4)), CInt(vDateBits(5)), 0)
				Txn.TxnDateValid = True
			End If
' cash withdrawals - special string for payee
			If vFields(fldTransCode) = "GM " Or vFields(fldTransCode) = "PK " Then
				Txn.Payee = "Kasopname"
			End If

			ConcatMemo Trim(vFields(fldOmschrijving1))
			ConcatMemo Trim(vFields(fldOmschrijving2))
			ConcatMemo Trim(vFields(fldOmschrijving3))
			ConcatMemo Trim(vFields(fldOmschrijving4))
			ConcatMemo Trim(vFields(fldOmschrijving5))
' very special checks!
			sTmp = Trim(vFields(fldOmschrijving1))
			If StartsWith(sTmp, "UWV ") _
			Or StartsWith(sTmp, "SPORT-EN") _
			Or StartsWith(sTmp, "BANKGIROLOTERIJ") Then
				Txn.Payee = sTmp
			ElseIf StartsWith(vFields(fldOmschrijving2), "STAATSLOTERIJ") Then
				Txn.Payee = "Staatsloterij" 
			ElseIf Instr(vFields(fldNaam), " BONI ") > 0 Then
				Txn.Payee = "Boni"
			ElseIf Instr(vFields(fldNaam), " HEMA ") > 0 Then
				Txn.Payee = "Hema"
			Else
			' direct debits almost always have the payee name in the last memo field
				If vFields(fldTransCode) = "IC " Then
					Txn.Payee = LastMemo
				End If
			End If
' speciaal voor onno
			If Txn.Payee = "GIROTEL/GIRONET ABONNEMENTSGELD" _
			Or Txn.Payee = "DEBETRENTE GIROKWARTAALKREDIET" Then
				Txn.Payee = "POSTBANK N.V."
			End If

' sort out a transaction ID - date, sequence are provided!
			Txn.CheckNum = CStr(Year(Txn.BookDate)) & _
				Right("0" & CStr(Month(Txn.BookDate)), 2) & _
				Right("0" & CStr(Day(Txn.BookDate)), 2) & _
				"." & vFields(fldTransNr)
		End If
	Loop
	LoadTextFile = True
End Function
