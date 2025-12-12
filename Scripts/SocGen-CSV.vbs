' MT2OFX Input Processing Script for Société Générale CSV format

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/SocGen-CSV.vbs 2     4-01-05 23:43 Colin $"

Const ScriptName = "Société Générale-CSV"
Const FormatName = "Société Générale CSV Format"
Const ParseErrorMessage = "Unable to parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

'Date;Nature de l'opération;Débit;Crédit;Monnaie;Date de valeur;Libellé interbancaire
' fld#	format		contents
'	1	dd/mm/yyyy	Date
Const	fldDate = 1
'	2	.*			Nature de l'opération
Const	fldTxnDate = 2
'	3	[\d ]*.\d\d	Débit
Const	fldAmtDebit = 3
'	4	[\d ]*.\d\d	Crédit
Const	fldAmtCredit = 4
'	5	XXX			Monnaie
Const	fldCurrency = 5
'	6	dd/mm/yyyy	Date de valeur
Const	fldValDate = 6
'	7	.*			Libellé interbancaire
Const	fldMemo = 7

Dim sTxnHeader
sTxnHeader = "Date;Nature de l'opération;Débit;Crédit;Monnaie;Date de valeur;Libellé interbancaire"

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
	Dim iYear, iMonth, iDay			' for dates, dd/mm/yyyy
	iYear = CInt(Mid(sDate,7,4))
	iMonth = CInt(Mid(sDate,4,2))
	iDay = CInt(Left(sDate,2))
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
End Sub

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
	Dim sLine
	RecogniseTextFile = False
	Dim vFields
	sLine = ReadLine()
	If Not StartsWith(sLine, "SG ") Then
		Exit Function
	End If
	If AtEOF() Then
		Exit Function
	End If
	sLine = ReadLine()
	If Not IsNumeric(Left(sLine, 5)) Then
		Exit Function
	End If
	Do While Not AtEOF()
		sLine = ReadLine()
' break out if we find our signature line
		If sLine = sTxnHeader Then
			Exit Do
		End If
' break out of function on first transaction line (not for us)
		If IsNumeric(Left(sLine, 1)) Then
			Exit Function
		End If
	Loop
	LogProgress ScriptName, "File Recognised"
	RecogniseTextFile = True
End Function

Function LoadTextFile()
	Dim sLine       ' holds a line
	Dim vFields     ' array of fields in the line
	Dim Stmt        ' holds the current statement
	Dim sTmp		' temporary String
	Dim bFirst		' true if txn is the first one

	LoadTextFile = False

	Set Stmt = NewStatement()
	Stmt.BankName = "SocGen"

' first loop processes header and stops after the transaction table header
	Do While Not AtEOF()
		sLine = ReadLine()
		If sLine = sTxnHeader Then
			Exit Do
		Elseif StartsWith(sLine, "Solde au;") Then
			sTmp = Trim(Mid(sLine,10))
			Stmt.ClosingBalance.BalDate = ParseDate(sTmp)
		Elseif StartsWith(sLine, "Solde;") Then
			vFields = ParseLineDelimited(sLine, ";")
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			If UBound(vFields) <> 3 Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			Stmt.ClosingBalance.Amt = ParseNumber(Replace(vFields(2)," ", ""), ",")
			Stmt.ClosingBalance.Ccy = Trim(vFields(3))
		Elseif IsNumeric(Left(sLine, 5)) Then
' second line looks like: "30003  00305  00020713745;EPC"
' fields are: bank code, branch code, account number;account name
			Stmt.BankName = Left(sLine, 5)
			Stmt.BranchName = Mid(sLine, 8, 5)
			Stmt.Acct = Mid(sLine, 15, 11)
		End If
	Loop

' this loop processes the actual transactions
	bFirst = True
	Do While Not AtEOF()
		sLine = ReadLine()
		If Len(sLine) > 0 Then
			If Left(sLine,1) = ";" Then
				ConcatMemo Trim(Mid(sLine,2))
			Else
				vFields = ParseLineDelimited(sLine, ";")
				If TypeName(vFields) <> "Variant()" Then
					MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
					Abort
					Exit Function
				End If
				If UBound(vFields) <> 7 Then
					MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
					Abort
					Exit Function
				End If
				NewTransaction
				Txn.Amt = ParseNumber(Replace(vFields(fldAmtCredit), " ", ""), ",") _
					+ ParseNumber(Replace(vFields(fldAmtDebit)," ",""), ",")
				Txn.ValueDate = ParseDate(vFields(fldValDate))
				Txn.BookDate = ParseDate(vFields(fldDate))
				Txn.IsReversal = False
				Txn.Payee = Trim(vFields(fldMemo))
				If Txn.Amt < 0 Then
					Txn.TxnType = "PAYMENT"
				Else
					Txn.TxnType = "DEP"
				End If
				Txn.Memo = vFields(fldMemo)
			
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
