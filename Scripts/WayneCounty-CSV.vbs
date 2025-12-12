' MT2OFX Input Processing Script for Wayne County Bank CSV format

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/WayneCounty-CSV.vbs 1     4-01-05 23:44 Colin $"

Const ScriptName = "WayneCounty-CSV"
Const FormatName = "Wayne County Bank CSV Format"
Const ParseErrorMessage = "Unable to parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

' fld#	format		contents
'	1	\d*			transaction number
Const	fldTxnNum = 1
'	2	mm/dd/yyyy	Date
Const	fldTxnDate = 2
'	3	.*			description
Const	fldDescription = 3
'	4	.*			memo (unused?)
Const	fldMemo = 4
'	5	-\d*.\d\d	amount debit
Const	fldAmtDebit = 5
'	6	\d*.\d\d	amount credit
Const	fldAmtCredit = 6
'	7	-?\d*.\d\d	balance (after this transaction)
Const	fldBalance = 7
'	8	\d*			check number
Const	fldCheckNum = 8
'	9	-?\d*.\d\d	fee (unused?)
Const	fldFee = 9

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
	Dim iYear, iMonth, iDay			' for dates, mm-dd-yyyy
	iYear = CInt(Mid(sDate,7,4))
	iDay = CInt(Mid(sDate,4,2))
	iMonth = CInt(Left(sDate,2))
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
	Do While Not AtEOF()
		sLine = ReadLine()
		If StartsWith(sLine, "Account Name: ") Then
			Exit Do
		End If
		If IsNumeric(Left(sLine,1)) Then
			Exit Function
		End If
		If StartsWith(sLine, "Transaction") Then
			vFields = ParseLineDelimited(sLine, ",")
			If TypeName(vFields) <> "Variant()" Then
				Exit Function
			End If
			If UBound(vFields) <> 9 Then
				Exit Function
			End If
			If vFields(fldTxnNum) <> "Transaction Number" _
				Or vFields(fldTxnDate) <> "Date" _
				Or vFields(fldDescription) <> "Description" _
				Or vFields(fldMemo) <> "Memo" _
				Or vFields(fldAmtDebit) <> "Amount Debit" _
				Or vFields(fldAmtCredit) <> "Amount Credit" _
				Or vFields(fldBalance) <> "Balance" _
				Or vFields(fldCheckNum) <> "Check Num" _
				Or vFields(fldFee) <> "Fee" Then
					Exit Function
			End If
		End If
	Loop
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
	sAcct = ""
	Do While Not AtEOF()
		sLine = ReadLine()
		If StartsWith(sLine, "Transaction Number") Then
			Exit Do
		Elseif StartsWith(sLine, "Account Number: ") Then
			sAcct = Trim(Mid(sLine,17))
		End If
	Loop
	bFirst = True
	
	Set Stmt = NewStatement()
	Stmt.Acct = Trim(sAcct)
	Stmt.BankName = "WayneCounty"
	Stmt.OpeningBalance.Ccy = "USD"
	Stmt.OpeningBalance.Amt = 0.0
	Stmt.ClosingBalance.Ccy = "USD"
	Stmt.ClosingBalance.Amt = 0.0

	Do While Not AtEOF()
		sLine = ReadLine()
		If Len(sLine) > 0 Then
			vFields = ParseLineDelimited(sLine, ",")
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			If UBound(vFields) <> 9 Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			NewTransaction
			Txn.Amt = ParseNumber(vFields(fldAmtCredit), ".") _
				+ ParseNumber(vFields(fldAmtDebit), ".") _
				+ ParseNumber(vFields(fldFee), ".")
			Txn.ValueDate = ParseDate(vFields(fldTxnDate))
			Txn.BookDate = Txn.ValueDate
			Txn.IsReversal = False
			Txn.Payee = Trim(vFields(fldDescription))
			If Txn.Amt < 0 Then
				Txn.TxnType = "PAYMENT"
			Else
				Txn.TxnType = "DEP"
			End If
			sTmp = Trim(vFields(fldCheckNum))
			If Len(sTmp) > 0 Then
				Txn.CheckNum = sTmp
			End If
			Txn.FITID = vFields(fldTxnNum)
			Txn.Memo = vFields(fldDescription)
			sTmp = Trim(vFields(fldMemo))
			If Len(sTmp) > 0 Then
				ConcatMemo sTmp
			End If
			
			Stmt.ClosingBalance.BalDate = Txn.BookDate				
			Stmt.ClosingBalance.Amt = ParseNumber(vFields(fldBalance), ".")
			If bFirst Then
				Stmt.OpeningBalance.BalDate = Txn.BookDate - 1
				Stmt.OpeningBalance.Amt = Stmt.ClosingBalance.Amt - Txn.Amt
				bFirst = False
			End If
		End If
	Loop
	LoadTextFile = True
End Function
