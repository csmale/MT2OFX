' MT2OFX Input Processing Script for Sydbank CSV format

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/Sydbank-CSV.vbs 2     08/02/2009 18:00 Pete $"

Const ScriptName = "Sydbank-CSV"
Const FormatName = "Sydbank CSV Format"
Const ParseErrorMessage = "Unable to parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

' "Dato";"Valør";"Tekst";"Beløb";"Saldo";"Afstemt"
' "31.12.2008";"31.12.2008";"Lønoverførsel";"-75,00";"21.461,20";"Udført"

' fld#	format		contents
'	1	dd.mm.yyyy	Date
Const	fldTxnDate = 1
'	2	dd.mm.yyyy	Date (not required) 
Const	fldValor = 2
'	3	.*			memo
Const	fldMemo = 3
'	4	-\d*,\d\d	amount
Const	fldAmount = 4
'	5	-?\d*.\d\d	balance (after this transaction)
Const	fldBalance = 5
'	6	.*			afstemt
Const	fldStatus = 6

' Property List is an array of arrays, each of which has the following elements:
'	1. Property key - used to reference properties
'	2. Property name - used as a label in the config screen
'	3. Property description - used as a description or tooltip in the config screen
'	4. Data type - ptString, ptBoolean, ptInteger, ptFloat, ptDate, ptChoice
'	5. Value list (will be displayed in a combobox) - array of values (Only with ptChoice)
Dim aPropertyList
aPropertyList = Array( _
	Array("AcctNum", "Account Number", _
		"The account number for Sydbank.", _
		ptString) _
	)

Sub Configure
	If ShowConfigDialog(ScriptName, aPropertyList) Then
		SaveProperties ScriptName, aPropertyList
	End If
End Sub

Sub Initialise()
    LogProgress ScriptName, "Initialise"
	If Not CheckVersion() Then
		Abort
	End If
	LoadProperties ScriptName, aPropertyList
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
	Dim iYear, iMonth, iDay			' for dates, dd.mm.yyyy
	iDay = CInt(Left(sDate,2))
	iMonth = CInt(Mid(sDate,4,2))
	iYear = CInt(Mid(sDate,7,4))
	ParseDate = DateSerial(iYear, iMonth, iDay)
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
	Dim sLine
	RecogniseTextFile = False
	Dim vFields
	If Not AtEOF() Then
		sLine = ReadLine()
		vFields = ParseLineDelimited(sLine, ";")
		If TypeName(vFields) <> "Variant()" Then
			Exit Function
		End If
		If UBound(vFields) <> 6 Then
			Exit Function
		End If
	' XX"Dato","Tekst","Beløb","Saldo","Status"
	' "Dato";"Valør";"Tekst";"Beløb";"Saldo";"Afstemt"
		If vFields(fldTxnDate) <> "Dato" _
			Or vFields(fldMemo) <> "Tekst" _
			Or vFields(fldAmount) <> "Beløb" _
			Or vFields(fldBalance) <> "Saldo" _
			Or vFields(fldStatus) <> "Afstemt" Then ' Afstemt changed from status
				Exit Function
		End If
	End If
	LogProgress ScriptName, "File Recognised"
	RecogniseTextFile = True
End Function

Function LoadTextFile()
	Dim sLine       ' holds a line
	Dim vFields     ' array of fields in the line
	Dim sAcct       ' last account number
	Dim Stmt        ' holds the current statement
	Dim sTmp		' temporary String
	Dim bFirst		' true if txn is the first one

	LoadTextFile = False
	
	sAcct = Trim(GetProperty("AcctNum"))
	If sAcct = "" Then
		MsgBox "Please set the Account Number in the script properties through the Options screen before converting a Sydbank file.",vbOkOnly,"Unknown account number"
		Abort
		LoadTextFile = False
		Exit Function
	End If

	bFirst = True
	
	Set Stmt = NewStatement()
	Stmt.Acct = sAcct
	Stmt.BankName = "SYBKDK22"
	Stmt.OpeningBalance.Ccy = "DKK"
	Stmt.OpeningBalance.Amt = 0.0
	Stmt.ClosingBalance.Ccy = "DKK"
	Stmt.ClosingBalance.Amt = 0.0

' eat column headers
	sLine = ReadLine()
	
	Do While Not AtEOF()
		sLine = ReadLine()
		If Len(sLine) > 0 Then
			vFields = ParseLineDelimited(sLine, ";")
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			If UBound(vFields) <> 6 Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			If vFields(fldStatus) = "nej" Then
				NewTransaction
				Txn.Amt = ParseNumber(vFields(fldAmount), ",")
				Txn.ValueDate = ParseDate(vFields(fldTxnDate))
				Txn.BookDate = Txn.ValueDate
				Txn.IsReversal = False
				Txn.Memo = Trim(vFields(fldMemo))
'XX				If StartsWith(Txn.Memo, "VDK ") Then
'XX					Txn.Payee = "Foreign Currency"
'XX				Else
					If IsNumeric(Left(Txn.Memo, 3)) And Mid(Txn.Memo,4,1) = " " Then
						Txn.Payee = Mid(Txn.Memo, 5)
					Else
						Txn.Payee = Txn.Memo
					End If
					If IsNumeric(Right(Txn.Payee, 5)) Then
						Txn.Payee = Trim(Left(Txn.Payee, Len(Txn.Payee)-5))
					End If
'XX				End If
				If Txn.Amt < 0 Then
					Txn.TxnType = "PAYMENT"
				Else
					Txn.TxnType = "DEP"
				End If
				
				Stmt.ClosingBalance.BalDate = Txn.BookDate				
				Stmt.ClosingBalance.Amt = ParseNumber(vFields(fldBalance), ",")
				If bFirst Then
					Stmt.OpeningBalance.BalDate = Txn.BookDate - 1
					Stmt.OpeningBalance.Amt = Stmt.ClosingBalance.Amt - Txn.Amt
					bFirst = False
				End If
			End If
		End If
	Loop
	LoadTextFile = True
End Function

