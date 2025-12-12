' MT2OFX Input Processing Script Bankinter (Spain) XLS format
Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/BankinterES-XLS.vbs 1     20/06/09 14:30 Colin $"

Const ScriptName = "BankinterES-XLS"
Const FormatName = "Bankinter (Spain) XLS"
Const ParseErrorMessage = "Cannot parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName
Const MinimumProgramVersion = "3"

Const DebugRecognition = False	' enables debug code in recognition
Const BankCode = "BKBKESMM"
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
Const CSVSeparator = ","
Const TxnLinePattern = ""
Const MinFieldsExpected = 5
Const MaxFieldsExpected = 5
Const DateSequence = "DMY"	' must be DMY, MDY, or YMD
Const DateSeparator = "-/. "	' can be empty for dates in e.g. "yyyymmdd" format
Const OldestLast = False		' True if transactions are in reverse order
Const InvertSign = False	' make credits into debits etc
Const CurrencyCode = "EUR"	' default if not specified in file
Const NoAvailableBalance = True		' True if file does not contain "Available Balance" information
Dim AccountNum: AccountNum = ""		' default if not specified in file
Dim BranchCode: BranchCode = ""		' default if not specified in file
Const SkipHeaderLines = 3	' number of lines to skip before the transaction data
Const ColumnHeadersPresent = True	' are the column headers in the file?
Const DecimalSeparator = ","	' as used in amounts
Const MemoChunkLength = 0	' if memo field consists of fixed length chunks
'DEL 10/04/08 19:34
Const TxnDatePattern = ""	' pattern to find transaction date in the memo
Dim TxnDateSequence: TxnDateSequence = Array(3,2,1,4,5,0)	' order of the info in the pattern: Y,M,D,H,M,S
Const PayeeLocation = 0		' start of payee in memo
Const PayeeLength = 0		' length of payee in memo
Dim MonthNames					' month names in dates
'MonthNames = Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
' Either give the month names in an array as above or use SetLocale to get the
' system strings for the given locale. Otherwise the default locale will be used.
' The MonthNames array must have a multiple of 12 elements, which run from Jan-Dec in groups of
' 12, i.e. "Jan".."Dec","January".."December" etc. Lower/upper case is not significant.
' SetLocale "nl-nl"

' Field name constants are now in MT2OFX.vbs
' For reference, they are:
' fldSkip, fldAccountNum, fldCurrency, fldClosingBal, fldAvailBal,
' fldBookDate, fldValueDate, fldAmtCredit, fldAmtDebit, fldMemo
' fldBalanceDate, fldAmount, fldPayee, fldTransactionDate, fldTransactionTime,
' fldChequeNum, fldCheckNum, fldFITID, fldEmpty, fldBranch

' Declare fields in the order they appear in the file as an array of arrays. The inner arrays
' contain a field ID from the list above followed by the exact column header.
' An optional third element in the inner arrays is used to contain a RegExp pattern for use instead of the
' literal text in the second element. If the pattern starts with "=", it is treated as a VBScript expression,
' where the characters "%1" are replaced with the contents of the field from the file.
' For example: "=Validate(""%1"")" would cause the function Validate to be called, which must return either
' True or False to indicate whether the validation passed.
' FECHA CONTABLE;FECHA VALOR;DESCRIPCION;IMPORTE;SALDO

Dim aFields
aFields = Array( _
	Array(fldBookDate, "FECHA CONTABLE"), _
	Array(fldValueDate, "FECHA VALOR"), _
	Array(fldMemo, "DESCRIPCION"), _
	Array(fldAmount, "IMPORTE"), _
	Array(fldClosingBal, "SALDO") _
)

' Dictionary to facilitate field lookup by field code
Dim FieldDict
Set FieldDict = CreateObject("Scripting.Dictionary")

' Property List is an array of arrays, each of which has the following elements:
'	1. Property key - used to reference properties
'	2. Property name - used as a label in the config screen
'	3. Property description - used as a description or tooltip in the config screen
'	4. Data type - ptString, ptBoolean, ptInteger, ptFloat, ptDate, ptChoose
'	5. Value list (will be displayed in a combobox) - array of values (Only with ptChoose)
'	6. Validation pattern (optional) - RegExp to validate the value entered
'		If the pattern starts with "=", the rest of the string is taken to be the name of a function in this
'		script which is called, with the value entered as a parameter, and which must return True if the value
'		is acceptable and False otherwise.
'	7. Validation error message (optional) - Message which will be displayed if the value entered fails the validation.
'		The script may instead define a function ValidationMessage which must return a string containing the message.
'		In both cases, "%1" in the string will be replaced by the value entered.
Dim aPropertyList: aPropertyList = Array()
'aPropertyList = Array( _
'	Array("AcctNum", "Account number", _
'		"The account number for " & ScriptName, _
'		ptString,,"=CheckAccount", "Please enter a valid account number.") _
'	)
Function CheckAccount(s)
	CheckAccount = False
	If Len(s) = 0 Then Exit Function
	CheckAccount = True
End Function

Sub Initialise()
    LogProgress ScriptVersion, "Initialise"
	If Not CheckVersion() Then
		Abort
	End If
' fill field lookup dictionary
' NB: only the last occurrence is remembered!
	Dim i
	For i=0 To UBound(aFields)
		FieldDict(aFields(i)(0)) = i+1
	Next
' Initialise dictionary of month names
	InitialiseMonths MonthNames
' get properties
	LoadProperties ScriptName, aPropertyList
End Sub

' function DescriptiveName
' returns a string with a descriptive name of this script
Function DescriptiveName()
	DescriptiveName = FormatName
End Function

Sub Configure
	If ShowConfigDialog(ScriptName, aPropertyList) Then
		SaveProperties ScriptName, aPropertyList
	End If
End Sub

Function ParseDate(sDate)
	ParseDate = ParseDateEx(sDate, DateSequence, DateSeparator)
End Function

' wrapper functions for XLS
Dim xExcel, xDocument, xSheet, xRow, ixRow
Set xExcel = Nothing
Function InitXLS()
	InitXLS = False
	If Ucase(Right(Session.InputFile.FileName, 4)) <> ".XLS" Then
		Exit Function
	End If
	ixRow = 0
	If Not (xExcel Is Nothing) Then
		InitXLS = True
		Exit Function
	End If
	On Error Resume Next
	Set xExcel = GetObject(,"Excel.Application")
	If xExcel Is Nothing Then
		Set xExcel = CreateObject("Excel.Application")		
	End If
	If xExcel Is Nothing Then
		Exit Function
	End If
	Set xDocument = xExcel.Workbooks.Open(Session.InputFile.FileName, False, True)
	If xDocument Is Nothing Then
		Set xExcel = Nothing
		Exit Function
	End If
	Set xSheet = xDocument.Worksheets(1)
	If xSheet Is Nothing Then
		xDocument.Close
		Set xDocument = Nothing
		Set xExcel = Nothing
		Exit Function
	End If
	InitXLS = True
End Function

Function ReadLineXLS()
	Dim vTmp(), iCol, sCol, dVal
	ReDim vTmp(MaxFieldsExpected)
	ixRow = ixRow + 1
	Set xRow = xSheet.Rows(ixRow)
	For iCol=1 To MaxFieldsExpected
		Select Case iCol
		Case 1, 2	' dates
			sCol = xRow.Cells(iCol).Value
		Case 4, 5	' amount
			sCol = CStr(xRow.Cells(iCol))
			If IsNumeric(sCol) Then
				dVal = CDbl(sCol)
				sCol = Replace(CStr(dVal), ".", ",")
			End If
		Case 3	' text
			sCol = xRow.Cells(iCol).Value
		End Select
		vTmp(iCol) = sCol
	Next
	ReadLineXLS = vTmp
End Function

Function AtEOFXLS()
	AtEOFXLS = ixRow > xSheet.UsedRange.Rows.Count
End Function

Sub CloseXLS()
	If Not (xDocument Is Nothing) Then
		xDocument.Close False
		Set xDocument = Nothing
	End If
	If Not (xExcel Is Nothing) Then
		Set xExcel = Nothing
	End If
End Sub

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
	Dim vFields
	Dim sLine
	Dim i
	Dim bTmp
	Dim sField
	Dim sPat
	RecogniseTextFile = False

	If Len(MinimumProgramVersion) > 0 Then
		If Not VersionAtLeast(MinimumProgramVersion) Then
			MsgBox "This MT2OFX script requires at least " & MinimumProgramVersion & " of the program and you have version " & Version & ".", _
				vbOKOnly+vbInformation, ScriptName
			Abort
			Exit Function
		End If
	End If

	If Not InitXLS() Then
		Exit Function
	End If
	For i=1 To SkipHeaderLines
		If AtEOFXLS() Then
			Exit Function
		End If
		vFields = ReadLineXLS()
	Next
	If AtEOFXLS() Then
		Exit Function
	End If
	vFields = ReadLineXLS()
	If TypeName(vFields) <> "Variant()" Then
		If DebugRecognition Then
			MsgBox "not var array",,ScriptName
		End If
		Exit Function
	End If
	If UBound(vFields) < MinFieldsExpected Or UBound(vFields) > MaxFieldsExpected Then
		If DebugRecognition Then
			MsgBox "Wrong number of fields - got " & UBound(vFields) & ", expected " _
			& MinFieldsExpected & "-" & MaxFieldsExpected & " - " & sLine,,ScriptName
		End If
		Exit function
	End If
	If ColumnHeadersPresent Then
		For i=1 To UBound(vFields)
			sField = Trim(vFields(i))
			If UBound(aFields(i-1)) > 1 Then
				If Left(aFields(i-1)(2), 1) = "=" Then
					sTmp = Replace(Mid(aFields(i-1)(2), 2), "%1", sField)
					If Not Eval(sTmp) Then
						Exit Function
					End If
				Else
					If Not StringMatches(sField, aFields(i-1)(2)) Then
						MsgBox "Field " & CStr(i) & ": '" & sField & "' does not match '" & aFields(i-1)(2) & "'",,ScriptName
						Exit Function
					End If
				End If
			Else
				If sField <> aFields(i-1)(1) Then
					If DebugRecognition Then
						MsgBox "Field " & CStr(i) & " " & sField & ", expecting " & aFields(i-1)(1),,ScriptName
					End If
					Exit Function
				End If
			End If
		Next
	Else
' pattern-match the first row
		For i=1 To UBound(vFields)
			sField = Trim(vFields(i))
			If UBound(aFields(i-1)) > 1 Then
				sPat = aFields(i-1)(2)
				If Left(sPat, 1) = "=" Then
					sTmp = Replace(Mid(sPat, 2), "%1", sField)
					bTmp = Eval(sTmp)
				Else
					bTmp = StringMatches(sField, sPat)
				End If
			Else
				Select Case aFields(i-1)(0)
				case fldSkip, fldMemo, fldPayee
					bTmp = True
				Case fldEmpty
					sPat = "(empty)"
					bTmp = (Len(sField) = 0)
				Case fldAccountNum
					sPat = "(account number)"
					bTmp = (Len(sField) > 0)
				Case fldBranch
					sPat = "(branch code)"
					bTmp = (Len(sField) > 0)
				case fldCurrency
					sPat = "[A-Z][A-Z][A-Z]"
					bTmp = StringMatches(sField, sPat)
				case fldClosingBal, fldAvailBal, fldAmtCredit, fldAmtDebit, fldAmount
					If DecimalSeparator = "." Then
						sPat = "[+-]?[ 0-9,]*(\.[0-9]*)?"
					Else
						sPat = "[+-]?[ 0-9\.]*(,[0-9]*)?"
					End If
					bTmp = StringMatches(sField, sPat)
				case fldBookDate, fldValueDate, fldTransactionDate, fldBalanceDate
	' NB: ParseDate will throw an error on an invalid date! need to sort this
					sPat = "(date)"
					bTmp = (ParseDate(sField) <> NODATE)
				Case fldTransactionTime
					sPat = "(time)"
					bTmp = (Len(sField) > 0)
				End Select
			End If
			If Not bTmp Then
				If DebugRecognition Then
					MsgBox "Field " & i & " (" & sField & ") failed to match '" & sPat & "'",,ScriptName
				End If
				Exit Function
			End If
		Next
	End If
	LogProgress ScriptName, "File Recognised"
	RecogniseTextFile = True
End Function

Function LoadTextFile()
	Dim sLine       ' holds a line
	Dim vFields     ' array of fields in the line
	Dim sAcct       ' last account number
	Dim sBranch		' branch code
	Dim Stmt        ' holds the current statement
	Dim sTmp		' temporary string
	Dim vDateBits	' parts of date
	Dim iSeq		' transaction sequence number
	Dim i
	Dim dBal		' temp balance date
	Dim sField		' field value being processed
	Dim dMaxDate	' latest txn/book date - if we don't have a statement date
	Dim dMinDate
	Dim dOpeningBal
	Dim FirstTxn

	LoadTextFile = False
'	sAcct = GetProperty("AcctNum")
'	If sAcct="" Then sAcct = AccountNum
	
	If Not InitXLS() Then
		Exit Function
	End If
	For i=1 To SkipHeaderLines
		vFields = ReadLineXLS()
		If StartsWith(vFields(1), "Número de cuenta: ") Then
			sAcct = Trim(Mid(vFields(1), 19))
		End If
	Next
	If ColumnHeadersPresent Then
		vFields = ReadLineXLS()
	End If
	Do While Not AtEOFXLS()
		vFields = ReadLineXLS()
		If TypeName(vFields) <> "Variant()" Then
			MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
			Abort
			Exit Function
		End If
		If UBound(vFields) < MinFieldsExpected Or UBound(vFields) > MaxFieldsExpected Then
			Message True, True, "Wrong number of fields - " & CStr(UBound(vFields)+1) & " - " & sLine, ScriptName
			Abort
			Exit function
		End If
		If Len(Trim(vFields(1))) = 0 Then
			dOpeningBal = vFields(5)
		Else
			If IsEmpty(Stmt) Then
				Set Stmt = NewStatement()
	' this initialisation should be in the class constructor!! (fixed in 3.3.5)
				Stmt.OpeningBalance.BalDate = NODATE
				Stmt.OpeningBalance.Ccy = CurrencyCode
				Stmt.OpeningBalance.Amt = dOpeningBal
				Stmt.AvailableBalance.BalDate = NODATE
				If Not NoAvailableBalance Then Stmt.AvailableBalance.Ccy = CurrencyCode
				Stmt.ClosingBalance.BalDate = NODATE				
				Stmt.ClosingBalance.Ccy = CurrencyCode
				iSeq = 0
				Stmt.BankName = BankCode
				Stmt.Acct = sAcct
				Stmt.BranchName = BranchCode
				dMaxDate = NODATE
				dMinDate = NODATE
				FirstTxn = True
			End If
			NewTransaction
			iSeq = iSeq + 1
			LastMemo = ""
			For i=1 To UBound(vFields)
				sField = Trim(vFields(i))
				Select Case aFields(i-1)(0)
				case fldSkip, fldEmpty
					' do nothing
				case fldAccountNum
					Stmt.Acct = sField
					sAcct = sField
				case fldBranch
					Stmt.BranchName = sField
					sBranch = sField
				case fldCurrency
					Stmt.OpeningBalance.Ccy = sField
					Stmt.ClosingBalance.Ccy = sField
					If Not NoAvailableBalance Then Stmt.AvailableBalance.Ccy = sField
				case fldClosingBal
					If OldestLast Then
						If FirstTxn Then
							Stmt.ClosingBalance.Amt = ParseNumber(sField, DecimalSeparator)
						End If
					Else
						Stmt.ClosingBalance.Amt = ParseNumber(sField, DecimalSeparator)
					End If
				case fldAvailBal
					Stmt.AvailableBalance.Amt = ParseNumber(sField, DecimalSeparator)
				case fldBookDate
					Txn.BookDate = vFields(i)
					If Txn.BookDate <> NODATE Then
						If dMaxDate = NODATE Or Txn.BookDate > dMaxDate Then
							dMaxDate = Txn.BookDate
						End If
						If dMinDate = NODATE Or Txn.BookDate < dMinDate Then
							dMinDate = Txn.BookDate
						End If
					End If
				case fldValueDate
					Txn.ValueDate = vFields(i)
				Case fldTransactionDate
					Txn.TxnDate = ParseDate(sField)
					Txn.TxnDateValid = (Txn.TxnDate <> NODATE)
					If Txn.TxnDate <> NODATE Then
						If dMaxDate = NODATE Or Txn.TxnDate > dMaxDate Then
							dMaxDate = Txn.TxnDate
						End If
						If dMinDate = NODATE Or Txn.TxnDate < dMinDate Then
							dMinDate = Txn.TxnDate
						End If
					End If
				Case fldTransactionTime
					If Txn.TxnDate <> NODATE And Len(sField)=5 Then
						Txn.TxnDate = Txn.TxnDate + TimeSerial(CInt(Left(sField,2)), _
							CInt(Mid(sField,4,2)),0)
					End If
				case fldAmtCredit
					Txn.Amt = Txn.Amt + Abs(ParseNumber(sField, DecimalSeparator))
				case fldAmtDebit
					Txn.Amt = Txn.Amt - Abs(ParseNumber(sField, DecimalSeparator))
				Case fldAmount
					Txn.Amt = ParseNumber(sField, DecimalSeparator)
				Case fldChequeNum
					Txn.CheckNum = sField
				case fldMemo
					ConcatMemo sField
				Case fldBalanceDate
					dBal = ParseDate(sField)
					If dBal > Stmt.ClosingBalance.BalDate Or Stmt.ClosingBalance.BalDate = NODATE Then
						Stmt.ClosingBalance.BalDate = dBal
						Stmt.AvailableBalance.BalDate = dBal
					End If
					If dBal < Stmt.OpeningBalance.BalDate Or Stmt.OpeningBalance.BalDate = NODATE Then
						Stmt.OpeningBalance.BalDate = dBal
					End If
				Case fldPayee
					If Len(sField) > 0 Then
						Txn.Payee = sField
					End If
				Case fldFITID
					Txn.FITID = sField
				End select
			Next
' correct the sign of the amount
			If vFields(5) = "-" Then
				Txn.Amt = -Txn.Amt
			End If

' transaction type
			If Txn.Amt < 0 Then
				Txn.TxnType = "PAYMENT"
			Else
				Txn.TxnType = "DEP"
			End If
			
			Dim sMemo
' find the payee, transaction type and txn date if we can
			sMemo = Txn.Memo			
			If Len(TxnDatePattern) > 0 Then
				Txn.TxnDate = TransDate(sMemo)
				If Txn.TxnDate <> NODATE Then Txn.TxnDateValid = True
			End If
						
' tidy up the memo
			If MemoChunkLength > 0 Then
				sMemo = Txn.Memo
				Txn.Memo = ""
				For i=1 To Len(sMemo) Step MemoChunkLength
					ConcatMemo Trim(Mid(sMemo, i, MemoChunkLength))
				Next
			End If

' keep tabs on the statement/balance Date
			Stmt.ClosingBalance.BalDate = dMaxDate
			Stmt.OpeningBalance.BalDate = dMinDate
			
			FirstTxn = False
		End If
	Loop
	LoadTextFile = True
End Function

Private Function TransDate(sMemo)
	Dim vDateBits
	Dim dTxn
	Dim iYear, iMonth, iDay, iHour, iMin, iSec
	dTxn = NODATE
	vDateBits = ParseLineFixed(sMemo, TxnDatePattern)
	If TypeName(vDateBits) = "Variant()" Then
		If TxnDateSequence(0) > 0 Then iYear = CInt(vDateBits(TxnDateSequence(0)))
		If TxnDateSequence(1) > 0 Then iMonth = CInt(vDateBits(TxnDateSequence(1)))
		If TxnDateSequence(2) > 0 Then iDay = CInt(vDateBits(TxnDateSequence(2)))
		If TxnDateSequence(3) > 0 Then iHour = CInt(vDateBits(TxnDateSequence(3)))
		If TxnDateSequence(4) > 0 Then iMin = CInt(vDateBits(TxnDateSequence(4)))
		If TxnDateSequence(5) > 0 Then iSec = CInt(vDateBits(TxnDateSequence(5)))
		If iYear < 100 Then iYear = iYear + 2000
		dTxn = DateSerial(iYear, iMonth, iDay) + TimeSerial(iHour, iMin, iSec)
	End If
	TransDate = dTxn
End Function
