' MT2OFX Input Processing Script for Rabobank NL MutAsc format

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/Rabo-MutAsc.vbs 10    25/11/08 22:23 Colin $"

Const ScriptName = "Rabo-MutAsc"
Const FormatName = "Rabobank Nederland MutAsc Formaat"
Dim NoOFXMessage : NoOFXMessage = "Van dit bestandstype kunt u geen OFC of OFX produceren omdat het geen saldoinformatie bevat." _
	& vbCrLf & vbCrLf & "Kies een ander uitvoerformaat zoals QIF."
Const NoOFXTitle = "Uitvoerformaat niet mogelijk."
Dim BadRecordTypeMessage : BadRecordTypeMessage = "Onbekend recordtype!!! " & vbCrLf & vbCrLf _
	& "Kan niet verder."
Dim BadRecordTypeTitle : BadRecordTypeTitle = ScriptName
Const ParseErrorMessage = "Kan regel niet ontleden."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName
Const Pattern2 = "(\d{10})([A-Z]{3})99999999992000     0     (.{10})(.{24})0(\d{13})([CD])(\d{6})(\d{6})000099999(.{16})99  "
Const Pattern3 = "(\d{10})[A-Z]{3}99999999993 {29} {3}(.{32})(.{32})(\d) {7}"
Const Pattern4 = "(\d{10})[A-Z]{3}99999999994(.{32})(.{32})(.{32}) {8}"

' Field numbering reflects only captures () in the above patterns!
' Fields for all record types
Const fldRekNr = 1
' Fields for record type 2
Const fld2Valuta = 2
Const fld2TegenRekNr = 3
Const fld2TegenRekNaam = 4
Const fld2Bedrag = 5
Const fld2CrDr = 6
Const fld2BoekDatum = 7
Const fld2ValutaDatum = 8
Const fld2Info = 9
' Fields for record type 3
Const fld3Omschrijving1 = 2
Const fld3Omschrijving2 = 3
Const fld3AantalVolgt = 4
' Fields for record type 4
Const fld4Omschrijving3 = 2
Const fld4Omschrijving4 = 3
Const fld4Omschrijving5 = 4

Sub Initialise()
    LogProgress ScriptName, "Initialise"
End Sub

' function DescriptiveName
' returns a string with a descriptive name of this script
Function DescriptiveName()
	DescriptiveName = FormatName
End Function

Function ParseDate(sDate)
	Dim iYear, iMonth, iDay	' for dates
	iYear = CInt(Left(sDate,2)) + 1900
	If iYear < 1970 Then
		iYear = iYear + 100
	End If
	iMonth = CInt(Mid(sDate,3,2))
	iDay = CInt(Mid(sDate,5,2))
	ParseDate = DateSerial(iYear, iMonth, iDay)
End Function

Function StartsWith(s, Prefix)
	StartsWith = (Left(s,Len(Prefix)) = Prefix)
End Function

Sub ConcatMemo(s)
	If s = "" Then
		Exit Sub
	End If
	If Len(Txn.FurtherInfo) > 0 Then
		Txn.FurtherInfo = Txn.FurtherInfo & Cfg.MemoDelimiter
	End If
	Txn.FurtherInfo = Txn.FurtherInfo & s
End Sub

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
	Dim sLine
	RecogniseTextFile = False
	sLine = ReadLine()
	If Len(sLine) <> 128 Then
'		MsgBox "Line len: " & CStr(Len(sLine))
		Exit Function
	End If
	If Mid(sLine, 11, 28) <> "EUR99999999992000     0     " Then
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
	Dim sTmp		' temp string
	Dim LineNum		' line number read from file

	LoadTextFile = False
'	If Session.OutputFileType = "OFX" Or Session.OutputFileType = "OFC" Then
'		MsgBox NoOFXMessage, vbOkOnly + vbCritical, NoOFXTitle
'		Abort
'		Exit Function
'	End If
	sAcct = ""
	LineNum = 0
	Do While Not AtEOF()
		sLine = Replace(ReadLine(), Chr(26), "")	' kill control-z at end of file
		LineNum = LineNum + 1
		If Len(sLine) > 0 Then
			sType = Mid(sLine, 24, 1)
			Select Case sType
			Case "2"
				sPat=Pattern2
			Case "3"
				sPat=Pattern3
			Case "4"
				sPat=Pattern4
			Case Else
				MsgBox "Regel " & CStr(LineNum) & ":" & vbCrLf & sLine, vbOkOnly+vbCritical, BadRecordTypeMessage
				Abort
				Exit Function
			End Select
			vFields = ParseLineFixed(sLine, sPat)
			If TypeName(vFields) <> "Variant()" Then
			MsgBox TypeName(vFields)
				MsgBox "Regel " & CStr(LineNum) & ":" & vbCrLf & sLine, vbOkOnly+vbCritical, ParseErrorMessage
				Abort
				Exit Function
			End If
	' set up new transaction, and start a new statement if the account # changes
			If sType="2" Then
				If sAcct <> vFields(fldRekNr) Then
					Set Stmt = NewStatement()
					sAcct = vFields(fldRekNr)
					Stmt.Acct = Trim(sAcct)
					Stmt.OpeningBalance.Ccy = vFields(fld2Valuta)
					Stmt.BankName = "Rabobank"
					Stmt.OpeningBalance.BalDate = ParseDate(vFields(fld2BoekDatum))
					Stmt.ClosingBalance.Ccy = ""	' to force 00000000 as date in ledger bal
				End If
				NewTransaction
			End If
			Select Case sType
			Case "2"
				Txn.Amt = CDbl(vFields(fld2Bedrag)) / 100.0
				If vFields(fld2CrDr) = "D" Then
					Txn.Amt = -Txn.Amt
				End If
				Txn.Payee = Trim(vFields(fld2TegenRekNaam))
				If Txn.Payee = "" Then
					Txn.Payee = "Onbekend"
				End If
				Txn.ValueDate = ParseDate(vFields(fld2ValutaDatum))
				Txn.BookDate = ParseDate(vFields(fld2BoekDatum))
				Txn.IsReversal = False
				Txn.TxnType = "OTHER"	' not much information in file
				ConcatMemo Trim(vFields(fld2Info))
				Stmt.ClosingBalance.BalDate = Txn.ValueDate
			Case "3"
				sTmp = Trim(vFields(fld3Omschrijving1))
				If StartsWith(sTmp, "Pinautomaat") Then
					Txn.TxnType = "POS"
				ElseIf StartsWith(sTmp, "Geldautomaat") Then
					Txn.TxnType = "ATM"
				End If
				ConcatMemo Trim(vFields(fld3Omschrijving1))
				ConcatMemo Trim(vFields(fld3Omschrijving2))
			Case "4"
				ConcatMemo Trim(vFields(fld4Omschrijving3))
				ConcatMemo Trim(vFields(fld4Omschrijving4))
				ConcatMemo Trim(vFields(fld4Omschrijving5))
			Case Else
				MsgBox "Internal error - bad record type not caught"
				Abort
				Exit Function
			End Select
		End If
	Loop
	If Cfg.TxnDumpFile <> "" Then
		DumpObjects Cfg.TxnDumpFile
	End If
	LoadTextFile = True
End Function
