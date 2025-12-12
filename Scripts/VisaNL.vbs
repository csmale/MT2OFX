' MT2OFX Input Processing Script for Visa NL via clipboard from website

' note about handling copy/paste from websites
' ============================================
' Different browsers do different things. When copying an HTML table to the clipboard,
' Mozilla Firefox conveniently separates cells by a Tab, whereas Internet
' Explorer just uses as space. It can become a complex problem to make the parsing
' handle both these variations. I have no idea at the moment how this will work
' with other browsers such as Opera.

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/VisaNL.vbs 6     19/04/08 23:18 Colin $"

Const ScriptName = "VisaNL"
Const FormatName = "Visa (NL) text from website - Rekeningoverzichten"
Const ParseErrorMessage = "Kan regel niet ontleden."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

Dim sPat
' 20-10-2004 21-10-2004 A RANDOM SHOP GB 73,49 / GBP € 107,95 
sPat = "(\d{2}-\d{2}-\d{4}) (\d{2}-\d{2}-\d{4}) (.+) (\d+,\d\d / [A-Z]{3})? (€ \d+,\d\d-?) "
Dim sBalPat
'   20-11-2004 Nieuw saldo   € 916,62 
sBalPat = "  (\d{2}-\d{2}-\d{4}) Nieuw saldo   (€ \d+,\d\d-?) "

' 5 transaction fields
	Dim sTxnDate
	Dim sBookDate
	Dim sPayee
	Dim sCurrency
	Dim sAmt


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

' Date format: DD-MM-YYYY
Function ParseDate(sDate)
	Dim iYear, iMonth, iDay			' for dates
	iDay = CInt(Left(sDate,2))
	iMonth = CInt(Mid(sDate,4,2))
	iYear = CInt(Mid(sDate,7,4))
	ParseDate = DateSerial(iYear, iMonth, iDay)
End Function

Function ParseAmount(sAmt)
' € 166,35-
	Dim sTmp
	sTmp = Mid(sAmt,3)
	If Right(sTmp,1) = "-" Then
		sTmp = Left(sTmp, Len(sTmp)-1)
' notice sign reversal as this is a credit card!
		ParseAmount = ParseNumber(sTmp, ",")
	Else
' notice sign reversal as this is a credit card!
		ParseAmount = -ParseNumber(sTmp, ",")
	End If
End Function

Sub ConcatMemo(s)
	If s = "" Then
		Exit Sub
	End If
	If Len(Txn.FurtherInfo) > 0 Then
		Txn.Memo = Txn.Memo & Cfg.MemoDelimiter
	End If
	Txn.Memo = Txn.Memo & s
End Sub

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
	Dim vFields
	Dim sLine
	Dim i
	RecogniseTextFile = False
' should start with 2-4 blank lines, then "On-line-accountoverzicht"
	For i=1 To 5
		If AtEOF() Then
			Exit Function
		End If
		sLine = Trim(Replace(ReadLine(), Chr(9), " "))
		If Len(sLine) <> 0 Then
			Exit For
		End If
	Next
	If sLine <> "On-line-accountoverzicht" Then
' MsgBox "acct ovz (line " & i & "):'" & sLine & "' - len=" & Len(sLine) & " char=0x" & Hex(AscW(Left(sLine, 1)))
		Exit Function
	End if
	
	RecogniseTextFile = True
	LogProgress ScriptName, "File Recognised - Visa NL"
End Function

Function LoadTextFile()
	Dim sLine       ' holds a line
	Dim vFields     ' array of fields in the line
	Dim sType       ' record type
	Dim sAcct       ' last account number
	Dim Stmt        ' holds the current statement
	Dim sTmp		' temporary String

	LoadTextFile = False
	sAcct = ""
	Set Stmt = NewStatement()
	Do While Not AtEOF()
		sLine = ReadLine()
		If Len(sLine) > 0 Then
			vFields = ParseLineDelimited(sLine, Chr(9))
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			If UBound(vFields) = 1 Then
				If Left(vFields(1), 14) = "Accountnummer " Then
					sAcct = Replace(Mid(vFields(1),15,11), " ", "")
					Stmt.Acct = Trim(sAcct)
					Stmt.BankName = "VisaNL"
					Stmt.OpeningBalance.Ccy = "EUR"
					Stmt.ClosingBalance.Ccy = "EUR"
				Elseif Mid(vFields(1), 14, 11) = "Nieuw saldo" Then
					sTmp = vFields(1)
					vFields = ParseLineFixed(sTmp, sBalPat)
					If TypeName(vFields) = "Variant()" Then
						Stmt.ClosingBalance.Amt = ParseAmount(vFields(2))
						Stmt.ClosingBalance.BalDate = ParseDate(vFields(1))
					End If		
				Elseif IsNumeric(Left(vFields(1),1)) Then
' transaction record!
					sTmp = vFields(1)
					vFields = ParseLineFixed(sTmp, sPat)
					If TypeName(vFields) <> "Variant()" Then
						MsgBox "'" & sTmp & "': " & TypeName(vFields)
						Abort
						Exit Function
					End If

					sAmt = vFields(5)
					sBookDate = vFields(1)
					sTxnDate = vFields(2)
					sPayee = vFields(3)
					sCurrency = vFields(4)
					DoTransaction
				End If
			Elseif UBound(vFields) = 5 Then
				If IsNumeric(Left(vFields(1),1)) Then
' transaction record!
					sAmt = vFields(5)
					sBookDate = vFields(1)
					sTxnDate = vFields(2)
					sPayee = vFields(3)
					sCurrency = vFields(4)
					DoTransaction
				Else
' non-transaction
					If Trim(vFields(3)) = "Nieuw saldo" Then
						Stmt.ClosingBalance.Amt = ParseAmount(vFields(5))
						Stmt.ClosingBalance.BalDate = ParseDate(vFields(2))
					End If
				End If
			End If
		End If
	Loop
	LoadTextFile = True
End Function

Function DoTransaction
	NewTransaction
	Txn.Amt = ParseAmount(sAmt)
	Txn.BookDate = ParseDate(sBookDate)
	Txn.ValueDate = Txn.BookDate
	Txn.TxnDate = ParseDate(sTxnDate)
	Txn.TxnDateValid = True
	Txn.Payee = Trim(sPayee)
	ConcatMemo Trim(sPayee)
	ConcatMemo Trim(sTxnDate)
	ConcatMemo Trim(sCurrency)
	If Txn.Amt < 0 Then
		Txn.TxnType = "PAYMENT"
	Else
		Txn.TxnType = "DEP"
	End If
End Function