' MT2OFX Input Processing Script Basic CSV format
' NB: This Script Will Not Work Without Customisation!

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/BanqueCantonaleVaudoiseCH-XLS.vbs 1     24/10/05 19:56 Colin $"

Const ScriptName = "BanqueCantonaleVaudoise-XLS"
Const FormatName = "Banque Cantonale Vaudoise XLS"
Const ParseErrorMessage = "Cannot parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

Const DebugRecognition = False	' enables debug code in recognition
Const BankCode = "BCVLCH2LXXX"
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
Const CSVSeparator = ","
Const TxnLinePattern = ""
Const NumFieldsExpected = 6
Const DateSequence = "DMY"	' must be DMY, MDY, or YMD
Const DateSeparator = "."	' can be empty for dates in e.g. "yyyymmdd" format
Const InvertSign = False	' make credits into debits etc
Dim CurrencyCode	' default if not specified in file
Dim AccountNum		' default if not specified in file
Const SkipHeaderLines = 10	' number of lines to skip before the transaction data
Const ColumnHeadersPresent = True	' are the column headers in the file?
Const DecimalSeparator = "."	' as used in amounts
Const MemoChunkLength = 0	' if memo field consists of fixed length chunks
Const TxnDatePattern = ".*(\d\d)\.(\d\d)\.(\d\d)\ (\d\d)\.(\d\d)"	' pattern to find transaction date in the memo
Const PayeeLocation = 0		' start of payee in memo
Const PayeeLength = 0		' length of payee in memo
Dim MonthNames					' month names in dates
'MonthNames = Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
' Either give the month names in an array as above or use SetLocale to get the
' system strings for the given locale. Otherwise the default locale will be used.
' The MonthNames array must have a multiple of 12 elements, which run from Jan-Dec in groups of
' 12, i.e. "Jan".."Dec","January".."December" etc. Lower/upper case is not significant.
' SetLocale "nl-nl"

Const fldSkip = 0
Const fldAccountNum = 1
Const fldCurrency = 2
Const fldClosingBal = 3
Const fldAvailBal = 4
Const fldBookDate = 5
Const fldValueDate = 6
Const fldAmtCredit = 7
Const fldAmtDebit = 8
Const fldMemo = 9
Const fldBalanceDate = 10
Const fldAmount = 11
Const fldPayee = 12
Const fldTransactionDate = 13
Const fldTransactionTime = 14
Const fldChequeNum = 15
Const fldFITID = 16

' Declare fields in the order they appear in the file as an array of arrays. The inner arrays
' contain a field ID from the list above followed by the exact column header.
' Date d'exécution;Opérations;Débit;Crédit;Date valeur;Solde
Dim aFields
aFields = Array( _
	Array(fldBookDate, "Date d'exécution"), _
	Array(fldMemo, "Opérations"), _
	Array(fldAmtDebit, "Débit"), _
	Array(fldAmtCredit, "Crédit"), _
	Array(fldValueDate, "Date valeur"), _
	Array(fldClosingBal, "Solde") _
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
Dim aPropertyList
aPropertyList = Array( _
	Array("AXACompte", "Numéro compte", _
		"Le numéro de compte pour AXA Belgique en format 000-000000-00.", _
		ptString) _
	)

' Special for BCV
Dim oExcelApp
Dim oWorkbook
Dim oSheet
Dim bIOpenedExcel
Dim iRow

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
' initialise some objects
	Set oExcelApp = Nothing
	Set oWorkbook = Nothing
	Set oSheet = Nothing
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

Function OpenExcel(sFile)
	bIOpenedExcel = False
	OpenExcel = False
	If oExcelApp Is Nothing Then
		On Error Resume Next
		Set oExcelApp = GetObject(, "Excel.Application")
		On Error Goto 0
	End If
	If oExcelApp Is Nothing Then
		Set oExcelApp = CreateObject("Excel.Application")
		bIOpenedExcel = True
	End If
	If oExcelApp Is Nothing Then
		MsgBox "Unable to create Excel.Application"
		Exit Function
	End If
	oExcelApp.Workbooks.Open sFile
	Set oWorkbook = oExcelApp.ActiveWorkbook
	Set oSheet = oWorkbook.Worksheets(1)
	OpenExcel = True
End Function

Sub CloseExcel
	If Not oSheet Is Nothing Then
		Set oSheet = Nothing
	End If
	If Not oWorkbook Is Nothing Then
		oWorkbook.Close False
		Set oWorkbook = Nothing
	End If
	If Not oExcelApp Is Nothing Then
		If bIOpenedExcel Then
			oExcelApp.Quit
		End If
		Set oExcelApp = Nothing
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
	If Not OpenExcel(Session.FileIn) Then
		Exit Function
	End If

	iRow = SkipHeaderLines + 1
	
	vFields = ReadRow()
	If TypeName(vFields) <> "Variant()" Then
		If DebugRecognition Then
			MsgBox "not var array"
		End If
		oExcelApp.Quit
		CloseExcel
		Exit Function
	End If
	If UBound(vFields) <> NumFieldsExpected Then
		If DebugRecognition Then
			MsgBox "wrong number of fields - got " & UBound(vFields) & ", expected " _
			& NumFieldsExpected & " - " & sLine
		End If
		CloseExcel
		Exit function
	End If
	If ColumnHeadersPresent Then
		For i=1 To NumFieldsExpected
			If vFields(i) <> aFields(i-1)(1) Then
				If DebugRecognition Then
					MsgBox "field " & CStr(i) & " " & aFields(i-1)(1) & " instead of " & vFields(i)
				End If
				CloseExcel
				Exit function
			End if
		Next
	Else
' pattern-match the first row
		For i=1 To NumFieldsExpected
			sField = Trim(vFields(i))
			If UBound(aFields(i-1)) > 2 Then
				bTmp = StringMatches(sField, aFields(i-1)(2))
			Else
				Select Case aFields(i-1)(0)
				case fldSkip, fldMemo, fldPayee
					bTmp = True
				Case fldAccountNum
					bTmp = (Len(sField) > 0)
				case fldCurrency
					bTmp = StringMatches(s, "[A-Z][A-Z][A-Z]")
				case fldClosingBal, fldAvailBal, fldAmtCredit, fldAmtDebit, fldAmount
					If DecimalSeparator = "." Then
						sPat = "[+-]?[ 0-9,]*(\.[0-9]*)?"
					Else
						sPat = "[+-]?[ 0-9\.]*(,[0-9]*)?"
					End If
					bTmp = StringMatches(sField, sPat)
				case fldBookDate, fldValueDate, fldTransactionDate, fldBalanceDate
	' NB: ParseDate will throw an error on an invalid date! need to sort this
					bTmp = (ParseDate(sField) <> NODATE)
				Case fldTransactionTime
					bTmp = (Len(sField) > 0)
				End Select
			End If
			If Not bTmp Then
				If DebugRecognition Then
					MsgBox "Field " & i & " (" & sField & ") failed to match"
				End If
				CloseExcel
				Exit Function
			End If
		Next
	End If
	LogProgress ScriptName, "File Recognised"
'	CloseExcel
	RecogniseTextFile = True
End Function

Function LoadTextFile()
	Dim sLine       ' holds a line
	Dim vFields     ' array of fields in the line
	Dim sAcct       ' last account number
	Dim Stmt        ' holds the current statement
	Dim sTmp		' temporary string
	Dim vDateBits	' parts of date
	Dim iSeq		' transaction sequence number
	Dim i
	Dim dBal		' temp balance date
	Dim sField		' field value being processed
	Dim dMaxDate	' latest txn/book date - if we don't have a statement date

	LoadTextFile = False
	If Not OpenExcel(Session.FileIn) Then
		Abort
		Exit Function
	End If

	sLine = Trim(Replace(CStr(oSheet.Cells(3,1).Value), Chr(160), " "))
	If StartsWith(sLine, "Compte : ") Then
		AccountNum = Trim(Mid(sLine, 10))
		Message False, True, "Account number found: " & AccountNum, ScriptName
	Else
		Message False, True, "Expected Compte, found " &sLine, ScriptName
	End If
	sLine = Trim(Replace(CStr(oSheet.Cells(7,1).Value), Chr(160), " "))
	If StartsWith(sLine, "Monnaie : ") Then
		CurrencyCode = Trim(Mid(sLine, 11))
		Message False, True, "Currency code found: " & CurrencyCode, ScriptName
	Else
		Message False, True, "Expected Monnaie, found " & sLine, ScriptName
	End If
		
	iRow = SkipHeaderLines + 1
	sAcct = ""
	If ColumnHeadersPresent Then
		iRow = iRow + 1
	End if
	Do While True
		vFields = ReadRow()
		If TypeName(vFields) <> "Variant()" Then
			MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
			Abort
			CloseExcel
			Exit Function
		End If
		If UBound(vFields) <> NumFieldsExpected Then
			Message True, True, "Wrong number of fields - " & CStr(UBound(vFields)+1) & " - " & sLine, ScriptName
			Abort
			CloseExcel
			Exit function
		End If
		If Len(Trim(vFields(1))) = 0 Then
			Exit Do
		End If
		If True Then
	' set up new transaction, and start a new statement if the account # changes
			If FieldDict.Exists(fldAccountNum) Then
				If sAcct <> vFields(FieldDict(fldAccountNum)) Then
					Set Stmt = NewStatement()
		' this initialisation should be in the class constructor!! (fixed in 3.3.5)
					Stmt.OpeningBalance.BalDate = NODATE
					Stmt.OpeningBalance.Ccy = CurrencyCode
					Stmt.AvailableBalance.BalDate = NODATE
					Stmt.AvailableBalance.Ccy = CurrencyCode
					Stmt.ClosingBalance.BalDate = NODATE				
					Stmt.ClosingBalance.Ccy = CurrencyCode
					iSeq = 0
					Stmt.BankName = BankCode
					dMaxDate = NODATE
				End If
			Else
				If IsEmpty(Stmt) Then
					Set Stmt = NewStatement()
		' this initialisation should be in the class constructor!! (fixed in 3.3.5)
					Stmt.OpeningBalance.BalDate = NODATE
					Stmt.OpeningBalance.Ccy = CurrencyCode
					Stmt.AvailableBalance.BalDate = NODATE
					Stmt.AvailableBalance.Ccy = CurrencyCode
					Stmt.ClosingBalance.BalDate = NODATE				
					Stmt.ClosingBalance.Ccy = CurrencyCode
					iSeq = 0
					Stmt.BankName = BankCode
					Stmt.Acct = AccountNum
					dMaxDate = NODATE
				End If
			End If
			NewTransaction
			iSeq = iSeq + 1
			LastMemo = ""
			For i=1 To UBound(vFields)
				sField = Trim(vFields(i))
				Select Case aFields(i-1)(0)
				case fldSkip
				case fldAccountNum
					Stmt.Acct = sField
					sAcct = sField
				case fldCurrency
					Stmt.OpeningBalance.Ccy = sField
					Stmt.ClosingBalance.Ccy = sField
					Stmt.AvailableBalance.Ccy = sField
				case fldClosingBal
' BCV: txns in reverse date order so closing balance is from first txn in file
					If Stmt.ClosingBalance.BalDate = NODATE Then
						Stmt.ClosingBalance.Amt = ParseNumber(sField, DecimalSeparator)
						Stmt.ClosingBalance.BalDate = Txn.BookDate
					End If
' BCV: opening balance is from last txn in file less last transaction
					Stmt.OpeningBalance.Amt = ParseNumber(sField, DecimalSeparator) - Txn.Amt
					Stmt.OpeningBalance.BalDate = Txn.BookDate
				case fldAvailBal
					Stmt.AvailableBalance.Amt = ParseNumber(sField, DecimalSeparator)
				case fldBookDate
					Txn.BookDate = ParseDate(sField)
					If Txn.BookDate <> NODATE Then
						If dMaxDate = NODATE Or Txn.BookDate > dMaxDate Then
							dMaxDate = Txn.BookDate
						End If
					End If
				case fldValueDate
					Txn.ValueDate = ParseDate(sField)
				Case fldTransactionDate
					Txn.TxnDate = ParseDate(sField)
					Txn.TxnDateValid = (Txn.TxnDate <> NODATE)
					If Txn.TxnDate <> NODATE Then
						If dMaxDate = NODATE Or Txn.TxnDate > dMaxDate Then
							dMaxDate = Txn.TxnDate
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
			If InvertSign Then
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
			If PayeeLocation > 0 And Len(Txn.Payee) = 0 Then
				Txn.Payee = Trim(Mid(sMemo, PayeeLocation, PayeeLength))
			End If
			If Len(TxnDatePattern) > 0 Then
				vDateBits = ParseLineFixed(Txn.Memo, TxnDatePattern)
				If TypeName(vDateBits) = "Variant()" Then
					Txn.TxnDate = DateSerial(Year(Stmt.OpeningBalance.BalDate), CInt(vDateBits(2)), CInt(vDateBits(1))) _
						+ TimeSerial(CInt(vDateBits(3)), CInt(vDateBits(4)), 0)
					Txn.TxnDateValid = True
				End If
			End If
						
' tidy up the memo
			If MemoChunkLength > 0 Then
				sMemo = Txn.Memo
				Txn.Memo = ""
				For i=1 To Len(sMemo) Step MemoChunkLength
					ConcatMemo Trim(Mid(sMemo, i, MemoChunkLength))
				Next
			End If

		End If
	Loop
	CloseExcel
	LoadTextFile = True
End Function

Function ReadRowEx(oXls, iRow, iCols)
	Dim xRow
	If iRow > oXls.Rows.Count Then
		ReadRowEx = "Row " & CStr(iRow) & " out of range"
MsgBox readrowex
		Exit Function
	End If
	Set xRow = oXls.Rows(iRow)
	Dim aRet()
	ReDim aRet(iCols)
	Dim i
	For i=1 To iCols
		aRet(i) = CStr(xRow.Cells(i).Value)
	Next
	ReadRowEx = aRet
	iRow = iRow + 1
End Function

Function ReadRow()
	ReadRow = ReadRowEx(oSheet, iRow, NumFieldsExpected)
End Function
