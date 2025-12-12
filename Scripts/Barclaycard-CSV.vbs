' MT2OFX Input Processing Script for Barclaycard CSV format

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/Barclaycard-CSV.vbs 1     20/03/05 23:38 Colin $"

Const ScriptName = "Barclaycard-CSV"
Const FormatName = "Barclaycard (UK) Comma-Separated Formaat"
Const ParseErrorMessage = "Unable to parse file."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

' Property List is an array of arrays, each of which has the following elements:
'	1. Property key - used to reference properties
'	2. Property name - used as a label in the config screen
'	3. Property description - used as a description or tooltip in the config screen
'	4. Data type - ptString, ptBoolean, ptInteger, ptFloat, ptDate, ptChoose
'	5. Value list (will be displayed in a combobox) - array of values (Only with ptChoose)
Dim aPropertyList
aPropertyList = Array( _
	Array("AcctNumber", "Account Number", _
		"Your Barclaycard Account Number. Please enter just the 16 digits, without any punctuation", _
		ptString) _
	)

Sub Configure
	If ShowConfigDialog(ScriptName, aPropertyList) Then
		SaveProperties ScriptName, aPropertyList
	End If
End Sub

' fld#	content
'	1	date
Const fldDate = 1
'	2	description
Const fldDescription = 2
'	3	unknown
Const fldUnknown = 3
'	4	amount
Const fldAmount = 4

Dim aMonths

Sub Initialise()
    LogProgress ScriptName, "Initialise"
	If Not CheckVersion() Then
		Abort
	End If
	LoadProperties ScriptName, aPropertyList
	Set aMonths = CreateObject("Scripting.Dictionary")
	aMonths("January") = 1
	aMonths("February") = 2
	aMonths("March") = 3
	aMonths("April") = 4
	aMonths("May") = 5
	aMonths("June") = 6
	aMonths("July") = 7
	aMonths("August") = 8
	aMonths("September") = 9
	aMonths("October") = 10
	aMonths("November") = 11
	aMonths("December") = 12
End Sub

' function DescriptiveName
' returns a string with a descriptive name of this script
Function DescriptiveName()
	DescriptiveName = FormatName
End Function

Function StartsWith(s, Prefix)
	StartsWith = (Left(s,Len(Prefix)) = Prefix)
End Function

' Barclaycard uses dd mmm yyyy in fldDate
Function ParseDate(sDate)
	Dim vArr
	vArr = Split(sDate, " ")
	ParseDate = NODATE
	If Not IsArray(vArr) Then
		Exit Function
	End If
	If UBound(vArr) <> 2 Then
		Exit Function
	End If
	If (Not IsNumeric(vArr(0))) Or (Not IsNumeric(vArr(2))) Or (Not aMonths.Exists(vArr(1))) Then
		Exit Function
	End If
	Dim iYear, iMonth, iDay			' for dates
	iDay = CInt(vArr(0))
	iMonth = CInt(aMonths(vArr(1)))
	iYear = CInt(vArr(2))
	ParseDate = DateSerial(iYear, iMonth, iDay)
End Function

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
	Dim vFields
	Dim sLine
	RecogniseTextFile = False
	sLine = ReadLine()
	vFields = ParseLineDelimited(sLine, ",")
	If TypeName(vFields) <> "Variant()" Then
		Exit Function
	End If
	If UBound(vFields) <> 4 Then
		Exit Function
	End If
	If ParseDate(vFields(fldDate)) = NODATE Then
		Exit Function
	End If
	If Trim(vFields(fldUnknown)) <> "" Then
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
	sAcct = Trim(GetProperty("AcctNumber"))
	If (Not IsNumeric(sAcct)) Or (Len(sAcct) <> 16) Then
		sAcct = ""
	End If
	If sAcct = "" Then
		MsgBox "Please set your Account Number in Options before converting a Barclaycard CSV file.",vbOkOnly,"Unknown account number"
		Configure
		sAcct = Trim(GetProperty("AcctNumber"))
		If (Not IsNumeric(sAcct)) Or (Len(sAcct) <> 16) Then
			sAcct = ""
		End If
		If sAcct = "" Then
			Abort
			LoadTextFile = False
			Exit Function
		End If
	End If

	iSeq = 0
	dSeq = NODATE
	Set Stmt = NewStatement()
	Stmt.Acct = sAcct
	Stmt.BankName = "Barclaycard"
	Stmt.OpeningBalance.Ccy = "GBP"	' no info in file
	Stmt.ClosingBalance.Ccy = ""	' to force 00000000 as date in ledger bal
	Do While Not AtEOF()
		sLine = ReadLine()
		If Len(sLine) > 0 Then
			vFields = ParseLineDelimited(sLine, ",")
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			If UBound(vFields) <> 4 Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			
' set up new transaction
			NewTransaction
			Txn.Amt = -ParseNumber(vFields(fldAmount), ".")
			Txn.ValueDate = ParseDate(vFields(fldDate))
			Txn.BookDate = ParseDate(vFields(fldDate))
			Txn.IsReversal = False
			Txn.Payee = Trim(vFields(fldDescription))

' default transaction type
			If Txn.Amt > 0 Then
				Txn.TxnType = "DEP"
			Else
				Txn.TxnType = "PAYMENT"
			End If

' sort out a transaction ID
			dBal = Txn.BookDate
			If dSeq = dBal Then
				iSeq = iSeq + 1
			Else
				iSeq = 1
				dSeq = dBal
			End If
			Txn.FITID = CStr(DatePart("yyyy", dSeq)) _
				& Right("00" & CStr(DatePart("y", dSeq)), 3) _
				& "." & CStr(iSeq)
		End If
	Loop
	LoadTextFile = True
End Function

