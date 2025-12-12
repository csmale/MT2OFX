' MT2OFX Input Processing Script Barry Rougeau's web capture format
Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/BarryRougeau.vbs 1     29/04/04 20:15 Colin $"

Const ScriptName = "BarryRougeau"
Const FormatName = "Barry Rougeau Text Capture Format"
Const ParseErrorMessage = "Unable to parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

' fld#	contents
'	1-3	date mm/dd/yyyy
Const fldMonth = 1
Const fldDay = 2
Const fldYear = 3
'	4	txn description
Const fldDescription = 4
'	5	txn amount
Const fldTxnAmt = 5
'	6	charge amount	
Const fldChargeAmt = 6
'	7	balance
Const fldBalance = 7

' pattern for transaction line
'Const TxnPat = "(\d{2})/(\d{2})/(\d{4}) (.*?) ([\-\d]\d*\.\d\d) ([\-\d]\d*\.\d\d) ([\-\d]\d*\.\d\d)"
Const TxnPat = "(\d{2})/(\d{2})/(\d{4}) *(.*?) *(\S*) *(\S*) *(\S*)"
' 04/24/2004 WITHDRAWAL - TX FLOWER MOUND 6400 MORRIS ROAD  -22.50 0.00 787.80

Sub Initialise()
    LogProgress ScriptVersion, "Initialise"
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
	sLine = ReadLine()
	If Trim(sLine) = "Transaction Date Transaction Description Transaction Amount Charge Amount Account Balance" Then
		LogProgress ScriptName, "File Recognised"
		RecogniseTextFile = True
	End If
End Function

Function LoadTextFile()
	Dim sLine       ' holds a line
	Dim vFields     ' array of fields in the line
' assume: all transactions on same account
' assume: transactions in reverse date order
' assume: map all transactions in one file to a single statement
' assume: all transactions for a given day in a single file (a day is not Split
'			over multiple files)
	Dim bDoneFirstTxn		' true if we have already done our first txn
	Dim Stmt	' statement object for this file
	Dim sDesc	' description text
	Dim iTmp	' temp for parsing description
	Dim iSeq	' sequence number for txn within a day
	Dim dThisDate	' date last processed
	
	LoadTextFile = False
	'	eat first (header) line
	If Not AtEof() Then
		sLine = ReadLine()
	End If

	Set Stmt = NewStatement()
	Stmt.Acct = "12345"
	Stmt.BankName = "BRougeau"
	Stmt.OpeningBalance.Ccy = "USD"
	Stmt.ClosingBalance.Ccy = "USD"

	Do While Not AtEOF()
		sLine = Trim(ReadLine())
		If Len(sLine) > 0 Then
			vFields = ParseLineFixed(sLine, TxnPat)
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			If Not bDoneFirstTxn Then
				Stmt.ClosingBalance.Amt = ParseNumber(vFields(fldBalance), ".")
				Stmt.ClosingBalance.BalDate = DateSerial(vFields(fldYear), vFields(fldMonth), vFields(fldDay))
				bDoneFirstTxn = True
			End If
			NewTransaction
			Txn.Amt = ParseNumber(vFields(fldTxnAmt), ".") _
				+ ParseNumber(vFields(fldChargeAmt), ".")
			Txn.BookDate = DateSerial(vFields(fldYear), vFields(fldMonth), vFields(fldDay))
			Txn.ValueDate = Txn.BookDate
	' keep count of transactions on this day
			If Txn.BookDate <> dThisDate Then
				dThisDate = Txn.BookDate
				iSeq = 1
			Else
				iSeq = iSeq + 1
			End If
			Txn.IsReversal = False
			Txn.BookingCode = ""
			Txn.Reference = ""
			Txn.BankReference = ""
			Txn.Payee = "Unknown"
			Txn.TxnDateValid = False
			Txn.TxnType = "OTHER"
' sort out a transaction ID - date, sequence are provided!
			Txn.FITID = CStr(Year(Txn.BookDate)) & _
				Right("0" & CStr(Month(Txn.BookDate)), 2) & _
				Right("0" & CStr(Day(Txn.BookDate)), 2) & _
				"." & CStr(iSeq)
' This is where the real work needs to be done in order to accommodate all
' formats of the transaction description. Three things can possibly be derived/inferred
' from this: the payee, the actual transaction date and the transaction type.
			sDesc = vFields(fldDescription)
			If StartsWith(sDesc, "WITHDRAWAL - ") Then
				Txn.TxnType = "ATM"
				Txn.Payee = "Withdrawal"
				sDesc = Mid(sDesc, 14)
			ElseIf StartsWith(sDesc, "EFT TRANS - ") Then
				Txn.TxnType = "PAYMENT"
				sDesc = Mid(sDesc, 13)
				iTmp = InStr(sDesc, ";")
				If iTmp > 0 Then
					Txn.Payee = Trim(Left(sDesc, iTmp-1))
				Else
					Txn.Payee = Trim(sDesc)
				End If
			Elseif StartsWith(sDesc, "POS PURCH - ") Then
				Txn.TxnType = "POS"
'				sDesc = Mid(sDesc, 13)
			Elseif StartsWith(sDesc, "SHR DRAFT ") Then
				Txn.TxnType = "CHECK"
				Txn.CheckNum = Mid(sDesc, 11)
'				sDesc = "Check " & Txn.CheckNum
			Elseif StartsWith(sDesc, "DIVIDEND ") Then
				Txn.TxnType = "DIV"
			Elseif StartsWith(sDesc, "CR CRD PMT ") Then
				Txn.TxnType = "PAYMENT"
			Elseif StartsWith(sDesc, "PYMT TRANS ") Then
				Txn.TxnType = "PAYMENT"
			Elseif StartsWith(sDesc, "TRANSFER ") Then
				Txn.TxnType = "XFER"
			Elseif StartsWith(sDesc, "JRNL ENTRY - ") Then
				If Txn.Amt >= 0.0 Then
					Txn.TxnType = "DEP"
				Else
					Txn.TxnType = "PAYMENT"
				End If
			End If
			Txn.FurtherInfo = sDesc
			
			' as the file is in reverse order, we can simply overwrite the
			' opening balance each time. The last values are from the earliest
			' transaction - but we must correct for the amount of that transaction
			Stmt.OpeningBalance.BalDate = Txn.BookDate
			Stmt.OpeningBalance.Amt = ParseNumber(vFields(fldBalance), ".") - Txn.Amt
		End If
	Loop
	LoadTextFile = True
End Function
