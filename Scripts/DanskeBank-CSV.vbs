' MT2OFX Input Processing Script for Danske Bank CSV format

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/DanskeBank-CSV.vbs 6     15/11/10 1:18 Colin $"

Const ScriptName = "DanskeBank-CSV"
Const FormatName = "Danske Bank CSV Formaat"
Const ParseErrorMessage = "Unable to parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

' "Dato","Tekst","Beløb","Saldo","Status"
' "14.01.2005","VDK SEK   106,00","-87,97","-249,11","Udført"

' fld#	format		contents
'	1	dd.mm.yyyy	Date
Const	fldTxnDate = 1
'	2	.*			memo
Const	fldMemo = 2
'	3	-\d*,\d\d	amount
Const	fldAmount = 3
'	4	-?\d*.\d\d	balance (after this transaction)
Const	fldBalance = 4
'	5	.*			status
Const	fldStatus = 5

' Property List is an array of arrays, each of which has the following elements:
'	1. Property key - used to reference properties
'	2. Property name - used as a label in the config screen
'	3. Property description - used as a description or tooltip in the config screen
'	4. Data type - ptString, ptBoolean, ptInteger, ptFloat, ptDate, ptChoice
'	5. Value list (will be displayed in a combobox) - array of values (Only with ptChoice)
Dim aPropertyList
aPropertyList = Array( _
	Array("AcctNum", "Account Number", _
		"The account number for Danske Bank.", _
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
		vFields = ParseLineDelimited(sLine, ",")
		If TypeName(vFields) <> "Variant()" Then
			Exit Function
		End If
' allow new column "Afstemt" (ignored)
		If UBound(vFields) < 5 Or UBound(vFields) > 6 Then
			Exit Function
		End If
	' "Dato","Tekst","Beløb","Saldo","Status"
	' "Dato","Tekst","Beløb","Saldo","Status","Afstemt"
		If vFields(fldTxnDate) <> "Dato" _
			Or vFields(fldMemo) <> "Tekst" _
			Or vFields(fldAmount) <> "Beløb" _
			Or vFields(fldBalance) <> "Saldo" _
			Or vFields(fldStatus) <> "Status" Then
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
		MsgBox "Please set the Account Number in the script properties through the Options screen before converting a Danske Bank file.",vbOkOnly,"Unknown account number"
		Abort
		LoadTextFile = False
		Exit Function
	End If

	bFirst = True
	
	Set Stmt = NewStatement()
	Stmt.Acct = sAcct
	Stmt.BankName = "DABADKKK"
	Stmt.OpeningBalance.Ccy = "DKK"
	Stmt.OpeningBalance.Amt = 0.0
	Stmt.ClosingBalance.Ccy = "DKK"
	Stmt.ClosingBalance.Amt = 0.0

' eat column headers
	sLine = ReadLine()
	
	Do While Not AtEOF()
		sLine = ReadLine()
		If Len(sLine) > 0 Then
			vFields = ParseLineDelimited(sLine, ",")
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			If UBound(vFields) < 5 Or UBound(vFields) > 6 Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			If vFields(fldStatus) = "Udført" Then
				NewTransaction
				Txn.Amt = ParseNumber(vFields(fldAmount), ",")
				Txn.ValueDate = ParseDate(vFields(fldTxnDate))
				Txn.BookDate = Txn.ValueDate
				Txn.IsReversal = False
				Txn.Memo = Trim(vFields(fldMemo))
				If StartsWith(Txn.Memo, "VDK ") Then
					Txn.Payee = "Foreign Currency"
				Else
					If IsNumeric(Left(Txn.Memo, 3)) And Mid(Txn.Memo,4,1) = " " Then
						Txn.Payee = Mid(Txn.Memo, 5)
					Else
						Txn.Payee = Txn.Memo
					End If
					If IsNumeric(Right(Txn.Payee, 5)) Then
						Txn.Payee = Trim(Left(Txn.Payee, Len(Txn.Payee)-5))
					End If
				End If
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
