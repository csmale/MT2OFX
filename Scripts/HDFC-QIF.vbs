' MT2OFX Input Processing Script to read QIF files
Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/HDFC-QIF.vbs 2     13/03/09 22:26 Colin $"

Const ScriptName = "HDFC-QIF"
Const FormatName = "HDFC Bank QIF Format"
Const ParseErrorMessage = "Unable to parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

' Customise This Lot!
Const csBankName = "HDFCINXX"
Dim sAccount: sAccount = "QIFImport"

' Property List is an array of arrays, each of which has the following elements:
'	1. Property key - used to reference properties
'	2. Property name - used as a label in the config screen
'	3. Property description - used as a description or tooltip in the config screen
'	4. Data type - ptString, ptBoolean, ptInteger, ptFloat, ptDate, ptChoice
'	5. Value list (will be displayed in a combobox) - array of values (Only with ptChoice)
Dim aPropertyList
aPropertyList = Array( _
	Array("QIFCurrency", "QIF Currency", _
		"The ISO currency code assumed for QIF input. Must be 3 letters, e.g. USD, EUR, GBP.", _
		ptCurrency), _
	Array("QIFDateFmt", "QIF Date Format", _
		"The date format in QIF input files.", _
		ptChoice, Array("DMY", "MDY")), _
	Array("QIFSwapPayeeMemo", "Swap payee and memo", _
		"If True, the memo field and the payee field are exchanged when the file is read.", _
		ptBoolean), _
	Array("QIFIntuitBankID", "Bank ID for Quicken", _
		"This value will be used as the value of INTU.BID in a QFX file. A value of 0 will not appear in the output.", _
		ptInteger), _
	Array("QIFNoTypeHeader", "Do not require '!Type:' header", _
		"If True, the file will not require a '!Type:' header and will be assumed to be '!Type:Bank' if none is found.", _
		ptBoolean), _
	Array("FITIDinMemo", "Memo field contains transaction ID", _
		"If True, the 'M' (memo) field is assumed to contain a unique transaction identifier.", _
		ptBoolean), _
	Array("IgnoreChequeNum", "Do not use the N-data as the cheque number", _
		"If True, the 'N' (cheque number) field is ignored. Otherwise it will be copied to the output.", _
		ptBoolean) _
	)

'Const csCurrency = "EUR"
'Const cbUSDates = True	' if True, dates in input are assumed to be MDY
Dim sCurrency	' currency from properties
Dim bUSDates	' if True, dates in input are assumed to be MDY
Dim sDateSequence	' derived from above - "MDY" or "DMY"
Dim bSwapPayeeMemo	' if True, the payee and memo fields are exchanged
Dim bNoType		' if True, no !Type: header will be expected.
Dim bFITIDinMemo	' if True, M-lines contain a FITID
Dim bIgnoreChequeNum	' If True, N-lines are ignored

Const ciMatchLines = 20
Dim MonthNames					' month names in dates
'MonthNames = Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
' Either give the month names in an array as above or use SetLocale to get the
' system strings for the given locale. Otherwise the default locale will be used.
' The MonthNames array must have a multiple of 12 elements, which run from Jan-Dec in groups of
' 12, i.e. "Jan".."Dec","January".."December" etc. Lower/upper case is not significant.
' SetLocale "nl-nl"

Sub Initialise()
    LogProgress ScriptName, "Initialise"
	If Not CheckVersion() Then
		Abort
	End If
	LoadProperties ScriptName, aPropertyList
' Initialise dictionary of month names
	InitialiseMonths MonthNames
End Sub

Sub Configure
	If ShowConfigDialog(ScriptName, aPropertyList) Then
		SaveProperties ScriptName, aPropertyList
	End If
End Sub

' function DescriptiveName
' returns a string with a descriptive name of this script
Function DescriptiveName()
	DescriptiveName = FormatName & ": Processes QIF files as input."
End Function

Function StartsWith(s, Prefix)
	StartsWith = (Left(s,Len(Prefix)) = Prefix)
End Function

' date formats:
' DD/MM/YY
' DD/MM/YYYY
' DD/MM'YYYY
' MM/DD/YY
' MM/DD/YYYY
' MM/DD'YYYY
' MM and DD might be a single digit, and might have a space prefix.
' what a mess.
Function ParseQIFDate(sDate)
	Dim iYear, iMonth, iDay			' for dates
	Dim sTmp
	Dim sDelim	' which delimiter? / or .?
	Dim iDelim
	ParseQIFDate = NODATE
	sDelim = "/"
	iDelim = InStr(sDate, sDelim)
	If iDelim=0 Then
		sDelim = "."
		iDelim = InStr(sDate, sDelim)
		If iDelim=0 Then
			sDelim = "-"
			iDelim = InStr(sDate, sDelim)
			If iDelim=0 Then
				Exit Function
			End If
		End If
	End If
	sTmp = Replace(sDate, "'", sDelim)
	ParseQIFDate = ParseDateEx(sTmp, sDateSequence, sDelim)
	If ParseQIFDate = NODATE Then
		MsgBox sTmp & "-" & ParseDateError,,"Error parsing date"
	End If
	Exit Function
End Function

Function ParseQIFAmount(sAmt)
	Dim sTmp
	Dim iPoint
	Dim iComma
	iPoint = InStr(samt, ".")
	iComma = InStr(sAmt, ",")
	If iPoint = 0 And iComma=0 Then
		ParseQIFAmount = CDbl(sAmt)
	ElseIf iPoint=0 And iComma>0 Then
		ParseQIFAmount = ParseNumber(sAmt, ",")
	Elseif iPoint>0 And iComma=0 Then
		ParseQIFAmount = ParseNumber(sAmt, ".")
	Elseif iPoint>iComma Then
		sTmp = Replace(sAmt, ",", "")
		ParseQIFAmount = ParseNumber(sTmp, ".")
	Else
		sTmp = Replace(sAmt, ".", "")
		ParseQIFAmount = ParseNumber(sTmp, ",")
	End If
End Function

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
	Dim i
	Dim sLine
	Dim re, m: Set re = New RegExp
	RecogniseTextFile = False
' HDFC Bank: extract account number from file name. If it fails, we don't recognise this file And
' someone else can have a try.
	re.Pattern = "(\d{14})-\d\d(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\.([Qq][Ii][Ff])"
	Set m = re.Execute(Session.InputFile.FileName)
	If m.Count = 0 Then
		Exit Function	' don't recognise this file name
	Else
		sAccount = m(0).SubMatches(0)
	End If

	bNoType = GetProperty("QIFNoTypeHeader")
	For i=1 To ciMatchLines	' try for a match in 20 lines
' eof and no match yet?
		If AtEOF() Then
			Exit Function
		End If
		sLine = ReadLine()
		If StartsWith(sLine, "!Type:Bank") Then
			' found bank statement
			sLine = ReadLine()
			If Mid(sLine,4,1) = "-" Then
'				MsgBox "Statement is from HDFC Bank",,"Bank recognised"
			End If
			Exit For
		Elseif StartsWith(sLine, "!Type:CCard") Then
			' found credit card statement
			Exit For
		Elseif StartsWith(sLine, "!Type:Cash") Then
			' found cash account statement
			Exit For
		Elseif sLine="^" Then
			' found end of data block before a valid !Type: assume bank if configured
			If bNoType Then
				Exit For
			End If
		End If
	Next
	If i > ciMatchLines Then
		Exit Function
	End If
	LogProgress ScriptName, "File Recognised - " & FormatName
	RecogniseTextFile = True
End Function

Function LoadTextFile()
	Dim sLine       ' holds a line
	Dim sAcct       ' last account number
	Dim dBalDate	' balance date
	Dim dBalAmt		' balance amount
	Dim Stmt        ' holds the current statement
	Dim sTmp		' temporary string
	Dim dBal		' temp date for check num
	Dim iSeq		' txn sequence num
	Dim dSeq		' date for sequence
	Dim bInStmt		' indicator - we are processing a statement
	Dim sType		' statement type
	Dim sRest		' rest of line
	Dim bNewTxn		' starting new transaction
	Dim bInAccount	' processing !Account
	Dim oSplit		' current split
	Dim bFirstLine	' true if processing the first line in the file
	Dim iAddrIdx	' next address index to be used

	LoadTextFile = False
'	MsgBox "HDFC script"
	If CStr(GetProperty("QIFIntuitBankID")) <> "0" Then
		Bcfg.IntuitBankID = CStr(GetProperty("QIFIntuitBankID"))
	End If
	sCurrency = "INR"
	sDateSequence = "MDY"
	bUSDates = True

	bSwapPayeeMemo = GetProperty("QIFSwapPayeeMemo")
	bNoType = GetProperty("QIFNoTypeHeader")
	bFITIDinMemo = GetProperty("FITIDinMemo")
	bIgnoreChequeNum = GetProperty("IgnoreChequeNum")
	
	sAcct = sAccount: dBalDate=NODATE: dBalAmt = 0
' if told not to expect a !Type header, need to check the first line
	bFirstLine = True
	bInStmt = False
	bInAccount = False
	Set oSplit = Nothing
	Do While Not AtEOF()
		sLine = ReadLine()
' handle missing !Type: line if configured
		If bFirstLine Then
			If bNoType And Left(sLine, 1) <> "!" Then
				sLine = "!Type:Bank"
				Rewind
			End If
		End If
		bFirstLine = False
		sType = Left(sLine, 1)
		sRest = Mid(sLine, 2)
		If bInAccount And sType<>"!" Then
			Select Case sType
			Case "N"	' account name/number
				sAcct = sRest
			Case "/"	' balance date
				dBalDate = ParseQIFDate(sRest)
			Case "$"	' balance amount
				dBalAmt = ParseQIFAmount(sRest)
			End Select
		ElseIf bInStmt And sType<>"!" Then
			If sType <> "^" And bNewTxn Then
				NewTransaction
				iAddrIdx = 0
'				Txn.SkipPayeeMapping = True
				bNewTxn = False
			End If
			Select Case sType
			Case "D"	' Date
				Txn.BookDate = ParseQIFDate(sRest)
' Opening/Closing Balance Dates are used for DTSTART and DTEND in OFX output
				If Stmt.OpeningBalance.BalDate = NODATE Or Stmt.OpeningBalance.BalDate > Txn.BookDate Then
					Stmt.OpeningBalance.BalDate = Txn.BookDate
				End If
				If Stmt.ClosingBalance.BalDate = NODATE Or Stmt.ClosingBalance.BalDate < Txn.BookDate Then
					Stmt.ClosingBalance.BalDate = Txn.BookDate
				End If
				Txn.ValueDate = Txn.BookDate
			Case "T"	' Amount
				Txn.Amt = ParseQIFAmount(sRest)
			Case "C"	' Cleared status
				Txn.ClearedStatus = sRest
			Case "N"	' Num (check or reference number)
				If Not bIgnoreChequeNum Then
					Txn.CheckNum = sRest
				End If
			Case "P"	' Payee
				Txn.Payee = sRest
			Case "M"	' Memo
				Txn.Memo = sRest
				If bFITIDinMemo Then
					Txn.FITID = sRest
				End If
			Case "A"	' Address (up to five lines; the sixth line is an optional message)
				iAddrIdx = iAddrIdx + 1
				Select Case iAddrIdx
				Case 1: Txn.Payee.Addr1 = sRest
				Case 2: Txn.Payee.Addr2 = sRest
				Case 3: Txn.Payee.Addr3 = sRest
				Case 4: Txn.Payee.Addr4 = sRest
				Case 5: Txn.Payee.Addr5 = sRest
				Case 6: Txn.Payee.Addr6 = sRest
				End Select
				' ignore for now
			Case "L"	' Category (Category/Subcategory/Transfer/Class)
				Txn.Category = sRest
			Case "S"	' Category in split (Category/Transfer/Class)
				If oSplit Is Nothing Then
					Set oSplit = Txn.Splits.AddNew
				End If
				oSplit.Category = sRest
			Case "E"	' Memo in split
				If oSplit Is Nothing Then
					Set oSplit = Txn.Splits.AddNew
				End If
				oSplit.Memo = sRest
			Case "$"	' Dollar amount of split
				If oSplit Is Nothing Then
					Set oSplit = Txn.Splits.AddNew
				End If
				oSplit.Amt = ParseQIFAmount(sRest)
' Amount must come last in split!
				Set oSplit = Nothing
' Note: Repeat the S, E, and $ lines as many times as needed for additional items in a split. If an item is omitted from the transaction in the QIF file, Quicken treats it as a blank item.
			Case "^"	' End of the entry
				If Txn.Amt < 0 Then
					Txn.TxnType = "PAYMENT"
				Else
					Txn.TxnType = "DEP"
				End If
' custom code for HDFC:
				If StartsWith(Txn.Payee, "POS ") Then
					Call ConcatMemo(Txn.Payee)
					Txn.Payee = Mid(Txn.Payee, 24)
					Txn.TxnType = "POS"
				End If
				If StartsWith(Txn.Payee, "ATW-") Then
					Call ConcatMemo(Txn.Payee)
					Txn.Payee = Mid(Txn.Payee, 25)
					Txn.TxnType = "POS"
				End If
' End custom code for HDFC
				If bSwapPayeeMemo Then
					sTmp = Txn.Memo
					Txn.Memo = Txn.Payee
					Txn.Payee = sTmp
				End If
				bNewTxn = True
			End Select
		Else
			bInStmt = False
			bInAccount = False
			If StartsWith(sLine, "!Type:") Then
' 20060111 CS: trim off extra blanks after !Type line
				sType = Trim(Mid(sLine, 7))
				Select Case sType
				Case "Bank", "CCard", "Cash"
					Set Stmt = NewStatement()
					Stmt.Acct = sAcct
					Stmt.QIFAcctType = sType
					If sType = "CCard" Then
						Stmt.AcctType = "CREDITCARD"
					Else
						Stmt.AcctType = "CHECKING"
					End If
					Stmt.BankName = csBankName
' 20060111 CS: always set currency!!!
					Stmt.OpeningBalance.Ccy = sCurrency
					If dBalDate <> NODATE Then
						Stmt.ClosingBalance.BalDate = dBalDate
						Stmt.ClosingBalance.Amt = dBalAmt
					Else
						Stmt.ClosingBalance.Ccy = ""	' to force 00000000 as date in ledger bal
					End If
					sAcct = sAccount: dBalDate=NODATE: dBalAmt = 0
					bInStmt = True
					bNewTxn = True
				End Select
			Elseif StartsWith(sLine, "!Account") Then
				bInAccount = True
			End If
		End If
	Loop
	LoadTextFile = True
End Function
