' MT2OFX Input Processing Script for Amex NL via clipboard from website

Option Explicit

' TODO:
' extract card number (XXXd out)
' extract balances and dates

Private Const ScriptVersion = "$Header: /MT2OFX/AmexNL-HTML.vbs 4     25/11/08 22:14 Colin $"

Const ScriptName = "AmexNL"
Const FormatName = "American Express (NL) text from website - Rekeningoverzichten"
Const ParseErrorMessage = "Kan regel niet ontleden."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

' the following lines wrap the big table with the transactions
Const csPatStart = "<!-- Start Transaction Details Content-->"
Const csPatEnd = "<!-- End Transaction Details Content -->"
            
Dim TxnTable
Const csPatRow = "<[Tt][Rr]>(.*?)</[Tt][Rr]>"
Const csPatCell = "\s*<[Tt][Dd][^>]*>(.*?)</[Tt][Dd]>\s*"
Const csGetText = "(?:</?.*>)*(.*)(?:/?<.*>)*"
Const csPatTag = "</?[^>]*>"

'Const csPatTxn = "<tr>\s*<td align=""left"" bgcolor=""#[0-9a-f]{6}""><font face=""arial"" size=""-1"">(.*)"

'<tr>
'            <td align="left" bgcolor="#ffeecc"><font face="arial" size="-1"> <b>03-01-08</b><br>
'            </font></td>
'            <td align="right" bgcolor="#ffeecc"><font face="arial" size="-1"> EUR 12,00</font></td>
'        </tr>
'        <tr>
'            <td colspan="2" align="left" bgcolor="#ffeecc"><font face="arial" size="-2"> P &amp; R PARKEER. SLOTERDIJK, AMSTERDAM<br>GOODS/SERVICES<br>12,00 EUROPEAN UNION EURO </font></td>
'        </tr>
        
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
' 166,35-
	Dim sTmp
	sTmp = Trim(sAmt)
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

Function GetTxnTable()
	Dim r: Set r = New Regexp
	Dim sText: sText = EntireFileHTML()
	If InStr(sText, csPatStart) = 0 Then
		GetTxnTable = ""
		Exit Function
	End If
'Dim fso: Set fso=CreateObject("Scripting.FileSystemObject"):Dim f: Set f=fso.OpenTextFile("c:\dump.txt", 2, True)
'f.Write sText:f.Close
	r.Global = False
	r.Multiline = True
	r.Pattern = ".*" & csPatStart & "(.*)" & csPatEnd & ".*"
' <!-- Start Transaction Details Content-->
' <!-- End Transaction Details Content -->
	Dim mm
	Set mm=r.Execute(sText)
' it works from a file, not from the clipboard???
' instr finds both...? might have something to do with line endings
	If mm.Count = 0 Then
		GetTxnTable = ""
	Else
'MsgBox "match: " & mm(0).Value
		GetTxnTable = Trim(mm(0).SubMatches(0))
	End If
'MsgBox GetTxnTable
End Function

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
	Dim sTmp
	Dim bRet
	sTmp = GetTxnTable()
	bRet = (Len(sTmp) > 0)
	If bRet Then
		LogProgress ScriptName, "File Recognised - Amex NL"
	End If
	RecogniseTextFile = bRet
End Function

Function LoadTextFile()
	Dim sLine       ' holds a line
	Dim vFields     ' array of fields in the line
	Dim sType       ' record type
	Dim sAcct       ' last account number
	Dim Stmt        ' holds the current statement
	Dim sTmp		' temporary String
	Dim sRow		' holds a row of the main table
	Dim sPayee		' payee
	Dim dTxnTotal	' total amount of transactions in this file
	Dim dLastDate	' date of previous txn
	Dim iDateSeq	' sequence number within date
	Dim iDelim
	Dim iCard
	Dim sTable
	Dim r: Set r = New regexp
	Dim rm, mm, m: Set rm = New Regexp
	Dim rc, cc, c: Set rc = New Regexp
	Dim rt, tt, t: Set rt = New Regexp

	LoadTextFile = False
	sAcct = ""
	Set Stmt = NewStatement()
	Stmt.AcctType = "CREDITCARD"
	Stmt.BankName = "AmexNL"
	Stmt.OpeningBalance.Ccy = "EUR"
	Stmt.ClosingBalance.Ccy = "EUR"
	dTxnTotal = 0
	
	dLastDate = NODATE
	sTable = GetTxnTable()
	r.Multiline = True
	r.Global = True
	r.pattern = csPatRow
	rm.Multiline = True
	rm.Global = True
	rm.Pattern = csPatCell
	rc.Multiline = True
	rc.Global = True
	rc.Pattern = csPatCell
	rt.Multiline = True
	rt.Global = True
	rt.Pattern = csPatTag
	
' get account number:             Rekening: XXXX-XXXXXX-71008<
Const csPatCard = "Rekening: ([\dX]{4}-[\dX]{5}-[\dX]{5})<"
	iCard = InStr(sTable, "Rekening: ")
	If iCard>0 Then
		iCard = iCard + 10
		iDelim = InStr(iCard, sTable, "<")
		Stmt.Acct = Trim(Mid(sTable, iCard, iDelim-iCard))
	End If
	
	Set mm = r.Execute(sTable)
	For Each m In mm
		sRow = Trim(replace(replace(m.SubMatches(0), vbCr, ""), vbLf, ""))
'		MsgBox "row: " & sRow
' a row can be either date/amount, description or rubbish
'            <TD align="left" bgColor="#FFEECC"><FONT face="arial" size="-1"> <B>29-01-08</B><BR>
		iDelim = InStr(sRow, ">")
		sTmp = Lcase(Left(sRow, iDelim))
		If InStr(sTmp, "align=""left""") > 0 And Instr(sTmp, "colspan=""2""") > 0 Then	' description
			Set cc = rc.Execute(sRow)
			If cc.Count = 1 Then
				Set c = cc(0)
				sTmp = Replace(c.Submatches(0), "<br>", Cfg.MemoDelimiter)
				sTmp = Replace(sTmp, "<BR>", Cfg.MemoDelimiter)
				sTmp = Trim(rt.Replace(sTmp, ""))
				Txn.Memo = HTMLDecode(sTmp)
				iDelim = InStr(Txn.Memo, Cfg.MemoDelimiter)
				If iDelim = 0 Then
					Txn.Payee = Txn.Memo
				Else
					Txn.Payee = Trim(Left(Txn.Memo, iDelim-1))
				End If
			End If
		ElseIf StartsWith(sTmp, "<td align=""left""") Then	' date/amount
			Set cc = rc.Execute(sRow)
			If cc.Count = 2 Then
				Set c = cc(0)
				sTmp = Trim(rt.Replace(c.Submatches(0), ""))
				If Not StartsWith(sTmp, "Totaal") Then
					NewTransaction
					Txn.BookDate = ParseDate(sTmp)
					Stmt.ClosingBalance.BalDate = Txn.BookDate
					If Stmt.OpeningBalance.BalDate = NODATE Then
						Stmt.OpeningBalance.BalDate = Txn.BookDate
					End If
					Set c = cc(1)
					sTmp = Trim(rt.Replace(c.Submatches(0), ""))
					Txn.Amt = ParseAmount(Replace(sTmp,"EUR", ""))
					If Txn.Amt < 0 Then
						Txn.TxnType = "PAYMENT"
					Else
						Txn.TxnType = "DEP"
					End If
					dTxnTotal = dTxnTotal + Txn.Amt
					If Txn.BookDate <> dLastDate Then
						iDateSeq = 1
						dLastDate = Txn.BookDate
					Else
						iDateSeq = iDateSeq + 1
					End If
					Txn.FITID = CStr(Year(dLastDate)) & "." & Right( "000" & CStr(DatePart("y", dLastDate)), 3) & "." & Right("000" & CStr(iDateSeq), 3)
				End If
			End If
		Else
'			MsgBox "row not relevant: " & sTmp
		End If
	Next
	Stmt.ClosingBalance.Amt = Stmt.OpeningBalance.Amt + dTxnTotal
	LoadTextFile = True
End Function


