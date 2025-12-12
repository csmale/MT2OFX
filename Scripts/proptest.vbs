' MT2OFX Input Processing Script for Argenta NL TSV format

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/proptest.vbs 3     2/11/05 23:03 Colin $"

Const ScriptName = "proptest"
Const FormatName = "test script for properties"
Const ParseErrorMessage = "Kan regel niet ontleden."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

' Property List is an array of arrays, each of which has the following elements:
'	1. Property key - used to reference properties
'	2. Property name - used as a label in the config screen
'	3. Property description - used as a description or tooltip in the config screen
'	4. Data type - ptString, ptBoolean, ptInteger, ptFloat, ptDate, ptChoose
'	5. Value list (will be displayed in a combobox) - array of values (Only with ptChoose)
Dim aPropertyList
dim ptcurrency: ptcurrency=7
aPropertyList = Array( _
	Array("myprop", "My Property", "This is a test property. Enter what you like", _
		ptString) , _
	Array("myprop2", "Choice Property", "This shows the value you just selected.", _
		ptString) , _
	Array("boolprop", "Another Property", "Here is another, just as a test", _
		ptBoolean) , _
	Array("dateprop", "Date Property", "Enter a date", _
		ptDate), _
	Array("numprop", "Number", "Pick a number...", _
		ptInteger), _
	Array("floatprop", "Floater", "Pick an amount...", _
		ptFloat), _
	Array("currprop", "Currency", "Pick a currency from the list", _
		ptCurrency), _
	Array("choiceprop", "Currency", "Pick a currency from the list", _
		ptChoice, Array("EUR", "USD", "GBP")) _
	)

' fld#	inhoud
'	1	rekeningnummer
Const fldRekNr = 1
'	2	rekeninghouder
'	3	leeg
'	4	onbekend (bedrag? saldo?)
'	5	valuta bij onbekend bedrag
'	6	saldodatum dd-mm-yyyy
Const fldSaldoDatum = 6
'	7	boekdatum	dd-mm-yyyy
Const fldBoekDatum = 7
'	8	tegenrekening
Const fldTgnRek = 8
'	9	begunstigde
Const fldNaam = 9
'	10	woonplaats begunstigde
Const fldWoonplaats = 10
'	11	bedrag
Const fldBedrag = 11
'	12	valuta
Const fldValuta = 12
'	13	txn code xx
Const fldTransCode = 13
'	14	omschrijving
Const fldOmschrijving = 14
'	15	onbekend

Sub Initialise()
    LogProgress ScriptName, "Initialise"
	If Not CheckVersion() Then
		Abort
	End If
	LoadProperties ScriptName, aPropertyList
End Sub

Sub Configure
	Dim sTmp
	sTmp = "def"
'	ChooseFromList cstr(sTmp), cstr(sTmp), cstr(sTmp), sTmp, True
	sTmp = ChooseFromList("abc\def\ghi", sTmp, "Test Title", "Pick a value", True)
MsgBox stmp
'ChooseFromList(sList As String, sDefault As String, sTitle As String, sHelp As String, bFixedList As Boolean) As String
	If sTmp <> "" Then
		SetProperty "myprop2", sTmp
	End If
	If ShowConfigDialog(ScriptName, aPropertyList) Then
		SaveProperties ScriptName, aPropertyList
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
	Dim iHour, iMin, iSec
	iDay = CInt(Left(sDate,2))
	iMonth = CInt(Mid(sDate,4,2))
	iYear = CInt(Mid(sDate,7,4))
	ParseDate = DateSerial(iYear, iMonth, iDay)
	If Len(sDate) > 10 Then
		iHour = CInt(Mid(sDate, 12, 2))
		iMin = CInt(Mid(sDate, 15, 2))
		iSec = CInt(Mid(sDate, 18, 2))
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
	vFields = ParseLineDelimited(sLine, vbTab)
	If TypeName(vFields) <> "Variant()" Then
		Exit Function
	End If
	If UBound(vFields) <> 15 Then
		Exit Function
	End If
	If vFields(3) <> "" Then
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

	LoadTextFile = False
	sAcct = ""
	Do While Not AtEOF()
		sLine = ReadLine()
		If Len(sLine) > 0 Then
			vFields = ParseLineDelimited(sLine, vbTab)
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			If UBound(vFields) <> 15 Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			
	' set up new transaction, and start a new statement if the account # changes
			If sAcct <> vFields(fldRekNr) Then
				Set Stmt = NewStatement()
				sAcct = vFields(fldRekNr)
				Stmt.Acct = Trim(sAcct)
				Stmt.BankName = "ArgentaNL"
				Stmt.OpeningBalance.Ccy = vFields(fldValuta)
				Stmt.OpeningBalance.BalDate = ParseDate(vFields(fldSaldoDatum))
				Stmt.ClosingBalance.Ccy = ""	' to force 00000000 as date in ledger bal
			End If
			NewTransaction
			Txn.Amt = ParseNumber(vFields(fldBedrag), ",")
			Txn.ValueDate = ParseDate(vFields(fldBoekDatum))
			Txn.BookDate = ParseDate(vFields(fldBoekDatum))
			Txn.Payee = Trim(vFields(fldNaam))
			Stmt.ClosingBalance.BalDate = Txn.ValueDate
			Txn.IsReversal = False
			Txn.FurtherInfo = Trim(vFields(fldOmschrijving))
			
' get transaction date out of description
			If StartsWith(Txn.FurtherInfo, "BETAALAUTOMAAT ") _
			Or StartsWith(Txn.FurtherInfo, "GELDAUTOMAAT ") Then
				Txn.TxnDate = ParseDate(Right(Txn.FurtherInfo, 19))
				Txn.TxnDateValid = True
			End If
			
' cash withdrawals - special string for payee
			If StartsWith(Txn.FurtherInfo, "GELDAUTOMAAT ") Then
				Txn.Payee = "Kasopname"
			Elseif StartsWith(Txn.FurtherInfo, "CHIPKNIP ") Then
				Txn.Payee = "Chipknip"
			Elseif StartsWith(Txn.FurtherInfo, "RENTE") Then
				Txn.Payee = "Rente"
			End If
			
' sometimes we don't get a payee!
			If Txn.Payee = "" Then
				If StartsWith(Txn.FurtherInfo, "BETAALAUTOMAAT ") Then
					sTmp = Mid(Txn.FurtherInfo, 16)	' lose BETAALAUTOMAAT
					sTmp = Left(sTmp, Len(sTmp)-19)	' lose date/time
					sTmp = Trim(Left(sTmp, 32))		' OFX: max len 32
					Txn.Payee = sTmp
				Else
					Txn.Payee = "Onbekend"
				End If
			End If

' transaction type
			If StartsWith(Txn.FurtherInfo, "GELDAUTOMAAT ") _
			Or StartsWith(Txn.FurtherInfo, "CHIPKNIP OP") Then
				Txn.TxnType = "ATM"
			Elseif StartsWith(Txn.FurtherInfo, "BETAALAUTOMAAT ") Then
				Txn.TxnType = "POS"
			Elseif StartsWith(Txn.FurtherInfo, "RENTE") Then
				Txn.TxnType = "INT"
			Elseif Txn.Amt > 0 Then
				Txn.TxnType = "DEP"
			Else
				Txn.TxnType = "PAYMENT"
			End If

' sort out a transaction ID
			dBal = ParseDate(vFields(fldSaldoDatum))
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

