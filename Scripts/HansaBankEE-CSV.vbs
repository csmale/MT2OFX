' MT2OFX Input Processing Script Hansapank (Swedbank) Estonia CSV format
' NB: THIS FILE MUST BE SAVED AS UTF-8

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/HansaBankEE-CSV.vbs 4     11/10/09 15:01 Colin $"

Dim Params: Set Params = New MT2OFXScript

With Params
	.MinimumProgramVersion = "3.5"
	.DebugRecognition = False	' enables debug code in recognition
	.ScriptName = "HansaBankEE-CSV"
	.FormatName = "Hansapank (Swedbank) Estonia CSV"
	.ParseErrorMessage = "Cannot parse line."
	.ParseErrorTitle = .ScriptName
	.BankCode = "HABAEE2X"
	.AccountNum = ""		' default if not specified in file
	.BranchCode = ""		' default if not specified in file
	.QuickenBankID = ""		' copied to INTU.BID if present
	.CurrencyCode = "EEK"	' default if not specified in file
	.ColumnHeadersPresent = True	' are the column headers in the file?
	.SkipHeaderLines = 0	' number of lines to skip before the transaction data
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
	.CSVSeparator = ","
	.DecimalSeparator = "."	' as used in amounts
	.TxnLinePattern = ""
	.DateSequence = "DMY"	' must be DMY, MDY, or YMD
	.DateSeparator = "-/. "	' can be empty for dates in e.g. "yyyymmdd" format
	.OldestLast = False		' True if transactions are in reverse order
	.InvertSign = False	' make credits into debits etc
	.NoAvailableBalance = True		' True if file does not contain "Available Balance" information
	.MemoChunkLength = 0	' if memo field consists of fixed length chunks
	.TxnDatePattern = ".*(\d\d)\.(\d\d)\.(\d{2}(\d{2})?)\ (\d\d):(\d\d).*"	' pattern to find transaction date in the memo
	.TxnDateSequence = Array(3,2,1,4,5,0)	' order of the info in the pattern (from 1 to 6): Y,M,D,H,M,S
	.PayeeLocation = 0		' start of payee in memo
	.PayeeLength = 0		' length of payee in memo
	.MonthNames = Empty
' "Client account","Row type","Date","Beneficiary/Payer","Details","Amount","Currency","Debit/Credit",
' "Transfer reference","Transaction type","Reference number","Document number",
' "221020217021","20","15.06.2007","","K:4797315010108240 14.06.2007 12:20 STATOIL PÄRNU HOMMIKU \\PÄRNU",
'  "1288.25","EEK","D","2007061500983947","K","",""
	.Fields = Array( _
		Array(fldAccountNum, "Client account"), _
		Array(fldSkip, "Row type"), _
		Array(fldBookDate, "Date"), _
		Array(fldPayee, "Beneficiary/Payer"), _
		Array(fldMemo, "Details"), _
		Array(fldAmount, "Amount"), _
		Array(fldCurrency, "Currency"), _
		Array(fldSkip, "Debit/Credit"), _
		Array(fldFITID, "Transfer reference"), _
		Array(fldSkip, "Transaction type"), _
		Array(fldSkip, "Reference number"), _
		Array(fldSkip, "Document number"), _
		Array(fldEmpty, "") _
	)
' min/max fields expected: default to size of Fields array. can be overridden here if required
	.MinFieldsExpected = 12
	.MaxFieldsExpected = 14
	.Properties = Array()

	Set .TransactionCallback = GetRef("TransactionCallback")
	Set .StatementCallback = GetRef("StatementCallback")
'	Set .HeaderCallback = GetRef("HeaderCallback")
'	Set .CustomDateCallback = GetRef("CustomDateCallback")
'	Set .CustomAmountCallback = GetRef("CustomAmountCallback")
	Set .PreParseCallback = GetRef("PreParseCallback")
	Set .ReadLineCallback = GetRef("ReadLineCallback")
'	Set .FinaliseCallback = GetRef("FinaliseCallback")
End With

' field indexes for standard layout. only needed for fields after "Beneficiary/Payer"
Dim nAmt: nAmt = 6
Dim nCcy: nCcy = 7
Dim nDebCred: nDebCred = 8
Dim nTxnType: nTxnType = 10


' "Client account","Row type","Date","Beneficiary/Payer","Beneficiary's account","Details","Amount","Currency","Debit/Credit",
' "Transfer reference","Transaction type","Reference number","Document number",
Dim Fields_EN2: Fields_EN2 = Array( _
		Array(fldAccountNum, "Client account"), _
		Array(fldSkip, "Row type"), _
		Array(fldBookDate, "Date"), _
		Array(fldPayee, "Beneficiary/Payer"), _
		Array(fldSkip, "Beneficiary's account"), _
		Array(fldMemo, "Details"), _
		Array(fldAmount, "Amount"), _
		Array(fldCurrency, "Currency"), _
		Array(fldSkip, "Debit/Credit"), _
		Array(fldFITID, "Transfer reference"), _
		Array(fldSkip, "Transaction type"), _
		Array(fldSkip, "Reference number"), _
		Array(fldSkip, "Document number"), _
		Array(fldEmpty, "") _
	)

' "Kliendi konto","Reatüüp","Kuupäev","Saaja/Maksja","Selgitus","Summa","Valuuta","Deebet/Kreedit",
' "Arhiveerimistunnus","Tehingu tüüp","Viitenumber","Dokumendi number",
Dim Fields_EE: Fields_EE = Array( _
		Array(fldAccountNum, "Kliendi konto"), _
		Array(fldSkip, "Reatüüp"), _
		Array(fldBookDate, "Kuupäev"), _
		Array(fldPayee, "Saaja/Maksja"), _
		Array(fldMemo, "Selgitus"), _
		Array(fldAmount, "Summa"), _
		Array(fldCurrency, "Valuuta"), _
		Array(fldSkip, "Deebet/Kreedit"), _
		Array(fldFITID, "Arhiveerimistunnus"), _
		Array(fldSkip, "Tehingu tüüp"), _
		Array(fldSkip, "Viitenumber"), _
		Array(fldSkip, "Dokumendi number"), _
		Array(fldEmpty, "") _
)
' "Kliendi konto","Reatüüp","Kuupäev","Saaja/Maksja","Selgitus","Summa","Valuuta","Deebet/Kreedit",
' "Arhiveerimistunnus","Tehingu tüüp","Viitenumber","Dokumendi number",
Dim Fields_EE2: Fields_EE2 = Array( _
		Array(fldAccountNum, "Kliendi konto"), _
		Array(fldSkip, "Reatüüp"), _
		Array(fldBookDate, "Kuupäev"), _
		Array(fldPayee, "Saaja/Maksja"), _
		Array(fldSkip, "Saaja konto"), _
		Array(fldMemo, "Selgitus"), _
		Array(fldAmount, "Summa"), _
		Array(fldCurrency, "Valuuta"), _
		Array(fldSkip, "Deebet/Kreedit"), _
		Array(fldFITID, "Arhiveerimistunnus"), _
		Array(fldSkip, "Tehingu tüüp"), _
		Array(fldSkip, "Viitenumber"), _
		Array(fldSkip, "Dokumendi number"), _
		Array(fldEmpty, "") _
)

' "Счëт клиентa","Строковый тип","Дата","Получатель/Плательщик","Пояснение","Сумма","Currency","Дебит/Кредит",
' "Архивный признак","Тип сделки","Номер ссылки","Номер документа",
Dim Fields_RU: Fields_RU = Array( _
		Array(fldAccountNum, "Счëт клиентa"), _
		Array(fldSkip, "Строковый тип"), _
		Array(fldBookDate, "Дата"), _
		Array(fldPayee, "Получатель/Плательщик"), _
		Array(fldMemo, "Пояснение"), _
		Array(fldAmount, "Сумма"), _
		Array(fldCurrency, "Currency"), _
		Array(fldSkip, "Дебит/Кредит"), _
		Array(fldFITID, "Архивный признак"), _
		Array(fldSkip, "Тип сделки"), _
		Array(fldSkip, "Номер ссылки"), _
		Array(fldSkip, "Номер документа"), _
		Array(fldEmpty, "") _
)
' "Счëт клиентa","Строковый тип","Дата","Получатель/Плательщик","Пояснение","Сумма","Currency","Дебит/Кредит",
' "Архивный признак","Тип сделки","Номер ссылки","Номер документа",
Dim Fields_RU2: Fields_RU2 = Array( _
		Array(fldAccountNum, "Счëт клиентa"), _
		Array(fldSkip, "Строковый тип"), _
		Array(fldBookDate, "Дата"), _
		Array(fldPayee, "Получатель/Плательщик"), _
		Array(fldSkip, "Счёт получателя"), _
		Array(fldMemo, "Пояснение"), _
		Array(fldAmount, "Сумма"), _
		Array(fldCurrency, "Currency"), _
		Array(fldSkip, "Дебит/Кредит"), _
		Array(fldFITID, "Архивный признак"), _
		Array(fldSkip, "Тип сделки"), _
		Array(fldSkip, "Номер ссылки"), _
		Array(fldSkip, "Номер документа"), _
		Array(fldEmpty, "") _
)


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
End Sub

' function DescriptiveName
' returns a string with a descriptive name of this script
Function DescriptiveName()
	DescriptiveName = FormatName
End Function

Sub Configure
	If ShowConfigDialog(Params.ScriptName, Params.Properties) Then
		SaveProperties Params.ScriptName, Params.Properties
	End If
End Sub

' single function to sort out languages and new/old formats
Sub GetFields()
    Dim bNewFormat: bNewFormat = False
    Session.InputFile.CodePage = CP_UTF8
    Dim sLine: sLine = ReadLine(): Call Rewind
    If InStr(sLine, ",""EEK"",") > 0 Then
        Params.ColumnHeadersPresent = False
        Session.InputFile.CodePage = 1257
    Else
        If StartsWith(sLine, """Kliendi konto"",") Then
            If InStr(sLine, Fields_EE2(4)(1)) > 0 Then
                Params.Fields = Fields_EE2
                bNewFormat = True
            Else
                Params.Fields = Fields_EE
            End If
        ElseIf StartsWith(sLine, """Счëт клиентa"",") Then
            If InStr(sLine, Fields_RU2(4)(1)) > 0 Then
                Params.Fields = Fields_RU2
                bNewFormat = True
            Else
                Params.Fields = Fields_RU
            End If
        Else
            If InStr(sLine, Fields_EN2(4)(1)) > 0 Then
                Params.Fields = Fields_EN2
                bNewFormat = True
            End If
        End If
	End If
   If bNewFormat Then
        nAmt = 7
        nCcy = 8
        nDebCred = 9
        nTxnType = 11
    End If
End Sub

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
    GetFields
    RecogniseTextFile = DefaultRecogniseTextFile(Params)
    If RecogniseTextFile Then
        LogProgress ScriptName, "File Recognised"
    End If
End Function

Function LoadTextFile()
    GetFields
    LoadTextFile = DefaultLoadTextFile(Params)
End Function

Dim dOpeningBal, dOpeningBalDate
Dim dClosingBal, dClosingBalDate
dOpeningBal = 0: dOpeningBalDate = NODATE
dClosingBal = 0: dClosingBalDate = NODATE

' callback functions, called from DefaultRecogniseTextFile and DefaultLoadTextFile
Function TransactionCallback(t, vFields)
'MsgBox "In transaction callback: " & t.Memo
	Dim iTmp
' fix up the sign
	If vFields(nDebCred) = "D" Then
		t.Amt = -t.Amt
		t.TxnType = "PAYMENT"
	End If
	t.Payee = TrimTrailingDigits(t.Payee)
	If StartsWith(t.Memo, "K:") Then
		t.Payee = Mid(t.Memo, 36)
		iTmp = InStr(t.Payee, "\")
		If iTmp>0 Then
			t.Payee = Trim(Left(t.Payee, iTmp-1))
		End If
	End If
	TransactionCallback = True
End Function

Function StatementCallback(s)
' MsgBox "In statement callback: " & s.StatementID
	s.OpeningBalance.Amt = dOpeningBal
	s.OpeningBalance.BalDate = dOpeningBalDate
	s.ClosingBalance.Amt = dClosingBal
	s.ClosingBalance.BalDate = dClosingBalDate
	dOpeningBal = 0
	dOpeningBalDate = NODATE
	dClosingBal = 0
	dClosingBalDate = NODATE
	StatementCallback = True
End Function

' PreParseCallback: returns True or False. True means the line can be processed; False means skip this line.
Function PreParseCallback(vFields)
'MsgBox "In preparse callback: " & UBound(vFields) & " fields."
	Dim dTmp
	PreParseCallback = False
' these are not transactions, but balances/totals
	Select Case vFields(nTxnType)
	Case "K2", "AI"	' transaction totals, estimated interest
		Exit Function
	Case "AS"	' opening balance
		dTmp = Params.ParseAmount(vFields(nAmt))
		If vFields(nDebCred) = "D" Then
			dTmp = -dTmp
		End If
		dOpeningBal = dTmp
		dOpeningBalDate = Params.ParseDate(vFields(3))
		Exit Function
	Case "LS"	' closing balance
		dTmp = Params.ParseAmount(vFields(nAmt))
		If vFields(nDebCred) = "D" Then
			dTmp = -dTmp
		End If
		dClosingBal = dTmp
		dClosingBalDate = Params.ParseDate(vFields(3))
		Exit Function
	End Select
' append currency code to account number because account can have multiple balances in different currencies
	vFields(1) = vFields(1) & vFields(nCcy)
	PreParseCallback = True
End Function

Function ReadLineCallback(sLine)
'MsgBox "In read line callback: " & sLine
' temp hack for codepage prob: u-tilde comes out as n-tilde
	sLine = Replace(sLine, ChrW(&h144), ChrW(&hF5))
' fix up \" at end of field
	ReadLineCallback = Replace(sLine, "\""", "\\""")
End Function
