' MT2OFX Input Processing Script to read Lloyds TSB QIF files
Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/LloydsTSB-QIF.vbs 3     25/11/08 22:23 Colin $"

Const ScriptName = "LloydsTSB-QIF"
Const FormatName = "Quicken Interchange Format - Lloyds TSB"
Const ParseErrorMessage = "Unable to parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

' Customise This Lot!
Const csBankName = "CTSBGB32"
Const csAccount = "QIFImport"

' Property List is an array of arrays, each of which has the following elements:
'	1. Property key - used to reference properties
'	2. Property name - used as a label in the config screen
'	3. Property description - used as a description or tooltip in the config screen
'	4. Data type - ptString, ptBoolean, ptInteger, ptFloat, ptDate, ptChoice
'	5. Value list (will be displayed in a combobox) - array of values (Only with ptChoice)
Dim aPropertyList
aPropertyList = Array( _
	Array("QIFAccount", "QIF Account", _
		"The Account Number assumed for QIF input. Only used if neither the QIF file nor its name does not contain an account number.", _
		ptString), _
	Array("QIFIntuitBankID", "Bank ID for Quicken", _
		"This value will be used as the value of INTU.BID in a QFX file. A value of 0 will not appear in the output.", _
		ptInteger) _
	)

Dim sCurrency	' currency from properties
Dim bUSDates	' if True, dates in input are assumed to be MDY
Dim sDateSequence	' derived from above - "MDY" or "DMY"
Dim bSwapPayeeMemo	' if True, the payee and memo fields are exchanged
Dim bNoType		' if True, no !Type: header will be expected.
Dim bFITIDinMemo	' if True, M-lines contain a FITID
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
	RecogniseTextFile = False
	bNoType = False
	For i=1 To ciMatchLines	' try for a match in 20 lines
' eof and no match yet?
		If AtEOF() Then
			Exit Function
		End If
		sLine = ReadLine()
		If StartsWith(sLine, "!Type:Bank") Then
			' found bank statement
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
	Dim sBankName	' bank name
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
	Dim iTmp		'
	Dim dTxn		' temp for transaction date

	LoadTextFile = False
	If CStr(GetProperty("QIFIntuitBankID")) <> "0" Then
		Bcfg.IntuitBankID = CStr(GetProperty("QIFIntuitBankID"))
	End If
	sCurrency = "GBP"
	sDateSequence = "DMY"
	If sDateSequence = "DMY" Then
		bUSDates = False
	Elseif sDateSequence = "MDY" Then
		bUSDates = True
	Else
		MsgBox "Please set the QIF Date Format in the script properties through the Options screen before converting a QIF file.",vbOkOnly,"Unknown date format"
		Abort
		LoadTextFile = False
		Exit Function
	End If
' file name example: 22615760_20060831_2310.qif
	iTmp = InStrRev(Session.FileIn, "\")
	sTmp = Mid(Session.FileIn, iTmp + 1)
	If StringMatches(sTmp, "\d{8}_20[0-9][0-9](0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])_([01][0-9]|2[0-3])[0-5][0-9]") Then
		sAcct = Left(sTmp, 8)
		Session.ServerTime = DateSerial( _
			CInt(Mid(sTmp, 10, 4)), _
			CInt(Mid(sTmp, 14, 2)), _
			CInt(Mid(sTmp, 16, 2))) _
			+ TimeSerial( _
			CInt(Mid(sTmp, 19, 2)), _
			CInt(Mid(sTmp, 21, 2)), _
			0)
		End If
' Account number: QIF contents overrides all this later if it is present
	If sAcct = "" Then
		sAcct = GetProperty("QIFAccount")
		If sAcct = "" Then
			sAcct = csAccount
		End If
	End If
' Bank name
	sBankName = csBankName
	bSwapPayeeMemo = False
	bNoType = False
	bFITIDinMemo = False
	dBalDate=NODATE: dBalAmt = 0
' if told not to expect a !Type header, need to check the first line
	bFirstLine = True
	bInStmt = False
	bInAccount = False
	Set oSplit = Nothing
	Do While Not AtEOF()
		sLine = Trim(ReadLine())
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
		If sType = "" Then
			' skip blank lines!
		ElseIf bInAccount And sType<>"!" Then
			Select Case sType
			Case "N"	' account name/number
				sAcct = sRest
			Case "/"	' balance date
				dBalDate = ParseQIFDate(sRest)
			Case "$"	' balance amount
				dBalAmt = ParseQIFAmount(sRest)
			End Select
		ElseIf bInStmt And sType<>"!" And sType<>"" Then
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
'				Txn.ValueDate = Txn.BookDate
			Case "T"	' Amount
				Txn.Amt = ParseQIFAmount(sRest)
			Case "C"	' Cleared status
				Txn.ClearedStatus = sRest
			Case "N"	' Num (check or reference number)
				Txn.CheckNum = sRest
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
				If bSwapPayeeMemo Then
					sTmp = Txn.Memo
					Txn.Memo = Txn.Payee
					Txn.Payee = sTmp
				End If
' special for Mike
				Txn.Memo = Txn.Payee
				iTmp = InStr(Txn.Payee, " . ")
				If iTmp > 0 Then
					Txn.Payee = Trim(Left(Txn.Payee, iTmp-1))
				End If
' atm withdrawals have a trans date at the end: 22JUL06
				sTmp = Right(Txn.Memo, 7)
' NB pattern runs out 31 december 2019
				If StringMatches(sTmp, "(0[1-9]|[12][0-9]|3[01])(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[01][0-9]") Then
					sTmp = Left(sTmp, 2) & "-" & Mid(sTmp, 3, 3) & "-" & Right(sTmp, 2)
					dTxn = ParseDateEx(sTmp, "DMY", "-")
					If dTxn <> NODATE Then
						Txn.TxnDate = dTxn
						Txn.TxnDateValid = True
					End If
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
					Stmt.BankName = sBankName
' 20060111 CS: always set currency!!!
					Stmt.OpeningBalance.Ccy = sCurrency
					If dBalDate <> NODATE Then
						Stmt.ClosingBalance.BalDate = dBalDate
						Stmt.ClosingBalance.Amt = dBalAmt
					Else
						Stmt.ClosingBalance.Ccy = ""	' to force 00000000 as date in ledger bal
					End If
					sAcct = csAccount: dBalDate=NODATE: dBalAmt = 0
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
