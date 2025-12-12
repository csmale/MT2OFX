' MT2OFX Input Processing Script to read QIF files
Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/Generic-QIF.vbs 24    9/01/10 23:21 Colin $"

Const ScriptName = "Generic-QIF"
Const FormatName = "Quicken Interchange Format"
Const ParseErrorMessage = "Unable to parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

' Customise This Lot!
Const csBankName = "QIFImport"
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
		"The Account Number assumed for QIF input. Only used if the QIF file does not contain an account number.", _
		ptString), _
	Array("QIFBank", "QIF Bank", _
		"The Bank Name assumed for QIF input. Only used if the QIF file does not contain an account number.", _
		ptString), _
	Array("QIFCurrency", "QIF Currency", _
		"The ISO currency code assumed for QIF input. Must be 3 letters, e.g. USD, EUR, GBP.", _
		ptCurrency), _
	Array("QIFDateFmt", "QIF Date Format", _
		"The date format in QIF input files.", _
		ptChoice, Array("DMY", "MDY", "YMD")), _
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
	Array("MemoPayeeLength", "Length of Payee at the start of the Memo", _
		"If no payee is specified for the transaction, the number of characters from the start of the Memo which are copied to the Payee", _
		ptInteger), _
	Array("BalDateIsStmtDate", "Use the latest transaction for the balance date", _
		"The closing balance date will be set to the date of the latest transaction instead of the invalid date '00000000'", _
		ptBoolean) _
	)

'Const csCurrency = "EUR"
Dim sCurrency	' currency from properties
Dim sDateSequence	' derived from above - "MDY" or "DMY" or "YMD"
Dim bSwapPayeeMemo	' if True, the payee and memo fields are exchanged
Dim bNoType		' if True, no !Type: header will be expected.
Dim bFITIDinMemo	' if True, M-lines contain a FITID
Dim bBalDateIsStmtDate	' if True, the closing balance date will be set to the latest transaction date
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
' 20070720 CS: remove UTF-8 euro sign found in ABN AMRO downloads!
	sTmp = Replace(sAmt, Chr(&hE2) & Chr(&h82) & Chr(&hAC), "")
	iPoint = InStr(sTmp, ".")
	iComma = InStr(sTmp, ",")
	If iPoint = 0 And iComma=0 Then
' 20070324 CS: catch non-numeric amounts (sometimes used in information-only pseudo-transactions)
		If IsNumeric(sTmp) Then
			ParseQIFAmount = CDbl(sTmp)
		Else
			ParseQIFAmount = 0
		End If
	ElseIf iPoint=0 And iComma>0 Then
		ParseQIFAmount = ParseNumber(sTmp, ",")
	Elseif iPoint>0 And iComma=0 Then
		ParseQIFAmount = ParseNumber(sTmp, ".")
	Elseif iPoint>iComma Then
		sTmp = Replace(sTmp, ",", "")
		ParseQIFAmount = ParseNumber(sTmp, ".")
	Else
		sTmp = Replace(sTmp, ".", "")
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
	bNoType = GetProperty("QIFNoTypeHeader")
	For i=1 To ciMatchLines	' try for a match in 20 lines
' eof and no match yet?
		If AtEOF() Then
			Exit Function
		End If
		sLine = LCase(ReadLine())
		If StartsWith(sLine, "!type:bank") Then
			' found bank statement
			Exit For
		Elseif StartsWith(sLine, "!type:ccard") Then
			' found credit card statement
			Exit For
		Elseif StartsWith(sLine, "!type:cash") Then
			' found cash account statement
			Exit For
		Elseif StartsWith(sLine, "!type:oth l") Then
			' found "Other" statement
			Exit For
		Elseif StartsWith(sLine, "!type:invst") Then
			' found investment statement - not supported yet!
			MsgBox "Input file contains investment transactions. MT2OFX cannot currently process this file.", _
				vbOKOnly+vbCritical, "QIF Investment Statement Detected"
			Abort
			Exit Function
			' do Nothing
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
	Dim MemoPayeeLength:MemoPayeeLength=GetProperty("MemoPayeeLength")
	If MemoPayeeLength < 0 Or MemoPayeeLength > 200 Then MemoPayeeLength=0

	LoadTextFile = False
	If CStr(GetProperty("QIFIntuitBankID")) <> "0" Then
		Bcfg.IntuitBankID = CStr(GetProperty("QIFIntuitBankID"))
	End If
	sCurrency = Ucase(GetProperty("QIFCurrency"))
	If Len(sCurrency) <> 3 Then
' 20080923 CS: Default to user's currency configured in Windows
		sCurrency = UserCurrency()
	End If
	If sCurrency = "" Then
		MsgBox "Please set the QIF Currency in the script properties through the Options screen before converting a QIF file.",vbOkOnly,"Unknown currency"
		Abort
		LoadTextFile = False
		Exit Function
	End If
	sDateSequence = UCase(GetProperty("QIFDateFmt"))
	If sDateSequence <> "DMY" And sDateSequence <> "MDY" And sDateSequence <> "YMD" Then
		MsgBox "Please set the QIF Date Format in the script properties through the Options screen before converting a QIF file.",vbOkOnly,"Unknown date format"
		Abort
		LoadTextFile = False
		Exit Function
	End If
' Account number: QIF contents overrides all this later if it is present
	sAcct = GetProperty("QIFAccount")
	If sAcct = "" Then
		sAcct = csAccount
	End If
' Bank name
	sBankName = GetProperty("QIFBank")
	If sBankName = "" Then
		sBankName = csBankName
	End If
	bSwapPayeeMemo = GetProperty("QIFSwapPayeeMemo")
	bNoType = GetProperty("QIFNoTypeHeader")
	bFITIDinMemo = GetProperty("FITIDinMemo")
	bBalDateIsStmtDate = GetProperty("BalDateIsStmtDate")
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
				If Len(Trim(Txn.Payee)) = 0 And MemoPayeeLength > 0 Then
					Txn.Payee = Trim(Left(Txn.Memo, MemoPayeeLength))
				End If
				bNewTxn = True
			End Select
		Else
			bInStmt = False
			bInAccount = False
			If StartsWith(LCase(sLine), "!type:") Then
' 20060111 CS: trim off extra blanks after !Type line
				sType = Trim(Mid(sLine, 7))
				Select Case LCase(sType)
				Case "bank", "ccard", "cash", "oth l"
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
					Stmt.ClosingBalance.Ccy = sCurrency
' 20070317 CS: use valid date in closing balance (bBalDateIsStmtDate)
					If (dBalDate <> NODATE) Or bBalDateIsStmtDate Then
						Stmt.ClosingBalance.BalDate = dBalDate
						Stmt.ClosingBalance.Amt = dBalAmt
					Else
						Stmt.ClosingBalance.Ccy = ""	' to force 00000000 as date in ledger bal
					End If
					sAcct = csAccount: dBalDate=NODATE: dBalAmt = 0
					bInStmt = True
					bNewTxn = True
				End Select
			Elseif StartsWith(LCase(sLine), "!account") Then
				bInAccount = True
			End If
		End If
	Loop
	LoadTextFile = True
End Function
