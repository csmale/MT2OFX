' MT2OFX Input Processing Script Basic CSV format
' NB: This Script Will Not Work Without Customisation!

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/Coop2GB-CSV.vbs 3     30/01/11 13:24 Colin $"

Dim dClosingBal

Dim Params: Set Params = New MT2OFXScript

With Params
	.MinimumProgramVersion = "3"
	.DebugRecognition = False	' enables debug code in recognition
	.ScriptName = "Coop2GB-CSV"
	.FormatName = "Cooperative Bank (2010 format) CSV"
	.ParseErrorMessage = "Cannot parse line."
	.ParseErrorTitle = .ScriptName
   .CodePage = 1252  ' Windows English / Western Europe
	.BankCode = "CPBKGB22"
	.AccountNum = "(in file)"		' default if not specified in file
	.BranchCode = ""		' default if not specified in file
	.AccountType = "CHECKING"	' can be CHECKING or CREDITCARD
	.QuickenBankID = ""		' copied to INTU.BID if present
	.CurrencyCode = "GBP"	' default if not specified in file
	.ColumnHeadersPresent = True	' are the column headers in the file?
	.SkipHeaderLines = 6	' number of lines to skip before the transaction data
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
	.CSVSeparator = ","
	.DecimalSeparator = "."	' as used in amounts
	.TxnLinePattern = ""
	.DateSequence = "DMY"	' must be DMY, MDY, or YMD
	.DateSeparator = "-/. "	' can be empty for dates in e.g. "yyyymmdd" format
	.OldestLast = True		' True if transactions are in reverse order
	.InvertSign = False	' make credits into debits etc
	.NoAvailableBalance = True		' True if file does not contain "Available Balance" information
	.MemoChunkLength = 0	' if memo field consists of fixed length chunks
	.TxnDatePattern = ".*(\d\d)\.(\d\d)\.(\d\d)\ (\d\d)\.(\d\d)"	' pattern to find transaction date in the memo
	.TxnDateSequence = Array(3,2,1,4,5,0)	' order of the info in the pattern (from 1 to 6): Y,M,D,H,M,S
	.PayeeLocation = 0		' start of payee in memo
	.PayeeLength = 0		' length of payee in memo
    .MonthNames = Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
' Date,Description,Bank        Reference,Customer Reference,   Credit,     Debit,    Additional Information,Running Balance     ,
' Date,Description,Bank     Reference,Customer  Reference,Credit,Debit,Additional Information,Running  Balance  
' 24/06/2010,Standing Order,IAN LANSDOWNE,,,2500.00,,,
    .Fields = Array( _
        Array(fldBookDate, "Date"), _
        Array(fldMemo, "Description"), _
        Array(fldMemo, "Bank Reference", "Bank +Reference"), _
        Array(fldMemo, "Customer Reference", "Customer +Reference"), _
        Array(fldAmtCredit, "Credit"), _
        Array(fldAmtDebit, "Debit"), _
        Array(fldMemo, "Additional Information"), _
        Array(fldClosingBal, "Running Balance", "Running +Balance"), _
        Array(fldSkip, "") _
    )
' min/max fields expected: default to size of Fields array. can be overridden here if required
	.MinFieldsExpected = 8
'	.MaxFieldsExpected = 1
'	.Properties = Array()
	Set .TransactionCallback = GetRef("TransactionCallback")
	Set .IsValidTxnLine = GetRef("IsValidTxnLine")
'	Set .PreParseCallback = GetRef("PreParseCallback")
	Set .HeaderCallback = GetRef("HeaderCallback")
	Set .StatementCallback = GetRef("StatementCallback")
'	Set .CustomDateCallback = GetRef("CustomDateCallback")
'	Set .CustomAmountCallback = GetRef("CustomAmountCallback")
'	Set .ReadLineCallback = GetRef("ReadLineCallback")
End With

' new format from november 2010
' ,    Date,,Description,Bank           Reference,,,,,Customer         Reference,,,Credit,,Debit,Additional Information,,Running  Balance  
' ,,04/11/2010,BACS Credit,5ZL224J86MKUL,,,,,PAYPAL TRANSFER,,,207.39,,,SERIAL NO - 609242,,2170.33
Dim vFieldsNew
vFieldsNew = Array( _
        Array(fldEmpty, ""), _
        Array(fldSkip, "Date"), _
        Array(fldBookDate, ""), _
        Array(fldMemo, "Description"), _
        Array(fldMemo, "Bank Reference", "Bank +Reference"), _
        Array(fldEmpty, ""), _
        Array(fldEmpty, ""), _
        Array(fldEmpty, ""), _
        Array(fldEmpty, ""), _
        Array(fldMemo, "Customer Reference", "Customer +Reference"), _
        Array(fldEmpty, ""), _
        Array(fldEmpty, ""), _        
        Array(fldAmtCredit, "Credit"), _
        Array(fldEmpty, ""), _
        Array(fldAmtDebit, "Debit"), _
        Array(fldMemo, "Additional Information"), _
        Array(fldEmpty, ""), _
        Array(fldClosingBal, "Running Balance", "Running +Balance") _
    )
Dim bNewFormat: bNewFormat = False

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

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
	RecogniseTextFile = DefaultRecogniseTextFile(Params)
	If RecogniseTextFile Then
		LogProgress ScriptName, "File Recognised"
	End If
End Function

Function LoadTextFile()
    LoadTextFile = DefaultLoadTextFile(Params)
End Function

' callback functions, called from DefaultRecogniseTextFile and DefaultLoadTextFile
' ths following implementations are functionally neutral or equivalent to the default processing
' in the class
Function HeaderCallback(sLine)
'MsgBox "In header callback: " & sLine
    Dim vFields, sTmp, iTmp
    
    If StartsWith(sLine, ",    Date,") Then
        Params.Fields = vFieldsNew
        bNewFormat = True
    End If
    vFields = ParseLineDelimited(sLine, Params.CSVSeparator)
    If StartsWith(sLine, "Account:") Then
        If UBound(vFields) > 1 Then
            sTmp = vFields(2)
            If Len(sTmp) = 0 Then sTmp = vFields(7)
            iTmp = InStr(sTmp, "-")
            If iTmp > 0 Then
                sTmp = Left(sTmp, iTmp-1)
            End If
            Params.AccountNum = Mid(sTmp, 7)
            Params.BranchCode = Left(sTmp, 6)
        End If
    ElseIf StartsWith(sLine, "Today's Cleared Balance:") Then
        sTmp = vFields(2)
        If Len(sTmp) = 0 Then sTmp = vFields(6)
        dClosingBal = Params.ParseAmount(sTmp)
' Today's Cleared Balance:,"3,060.66",,,,,,,
    End If
    HeaderCallback = True
End Function
Function TransactionCallback(t, vFields)
    Dim iTmp, sType, sPayee, sPayee2
    ' MsgBox "In transaction callback: " & t.Memo
    If bNewFormat Then
        sType = vFields(4)
        sPayee = vFields(10)
        sPayee2 = vFields(5)
    Else
        sType = vFields(2)
        sPayee = vFields(4)
        sPayee2 = vFields(3)
    End If
    Select Case sType
    Case "Cheque"
        t.TxnType = "CHECK"
        t.CheckNum = sPayee
    Case "Direct Debit"
        t.TxnType = "DIRECTDEBIT"
        t.Payee = sPayee
        If Len(t.Payee) = 0 Then t.Payee = vFields(3)
    Case "First Direct Debit"
        t.TxnType = "DIRECTDEBIT"
        t.Payee = sPayee
    Case "Faster Payment"
        t.Payee = sPayee
    Case "Purchase"
        t.TxnType = "POS"
        t.Payee = Trim(Mid(sPayee, 5))
    Case "Refund"
        t.TxnType = "POS"
        t.Payee = Trim(Mid(sPayee, 5))
    Case "Standing Order"
        t.Payee = sPayee
        If Len(t.Payee) = 0 Then t.Payee = sPayee2
    Case "BACS Credit"
        t.Payee = sPayee
        If Len(t.Payee) = 0 Then t.Payee = sPayee2
    Case "ATM Debit"
        t.TxnType = "ATM"
        'vFields(4) is like 5038LINK18:08AUG24
        GetTxnDate t, sPayee
    Case "Bank Credit Interest"
        t.TxnType = "INT"
    End Select
    TransactionCallback = True
End Function
Function CustomDateCallback(sDate)
MsgBox "In custom date callback: " & sDate
	CustomDateCallback = ParseDateEx(sDate, Params.DateSequence, Params.DateSeparator)
End Function
Function CustomAmountCallback(sAmt)
MsgBox "In custom amount callback: " & sAmt
	CustomAmountCallback = ParseNumber(sAmt, Params.DecimalSeparator)
End Function
Function ReadLineCallback(sLine)
MsgBox "In read line callback: " & sLine
	ReadLineCallback = sLine
End Function
' PreParseCallback: returns True or False. True means the line can be processed; False means skip this line.
Function PreParseCallback(vFields)
MsgBox "In preparse callback: " & UBound(vFields) & " fields."
    PreParseCallback = True
End Function
Sub StatementCallback(Stmt)
'MsgBox "In statement callback"
    Stmt.ClosingBalance.Amt = dClosingBal
End Sub
Function FinaliseCallback()
MsgBox "In finalisation callback"
	FinaliseCallback = True
End Function
' IsValidTxnLine can return:
' txnlineSKIP: skip this line
' txnlineNORMAL: this is the first line of a new transaction
' txnlineCONTINUATION: this line continues the previous transaction
' if using continuation lines, set Params.Fields and Params.TxnLinePattern before returning!
Function IsValidTxnLine(sLine)
' MsgBox "In IsValidTxnLine callback: " & sLine
    Dim sTmp
    If bNewFormat Then
        If Left(sLine, 2) = ",," Then
            sTmp = Mid(sLine, 3, 1)
        Else
            IsValidTxnLine = txnlineSKIP
            Exit Function
        End If
    Else
        sTmp = Left(sLine, 1)
    End If
    If IsNumeric(sTmp) Then
        IsValidTxnLine = txnlineNORMAL
    Else
        IsValidTxnLine = txnlineSKIP
    End If
End Function

' sDate is like 5038LINK18:08AUG24
Function GetTxnDate(t, ByVal sDate)
    Dim iTmp, iDay, iMon, iYear, iHour, iMin, dTmp, sMon
    iTmp = InStr(sDate, "LINK")
    If iTmp > 0 Then
        sDate = Mid(sDate, iTmp+4)
        iHour = CInt(Left(sDate, 2))
        iMin = CInt(Mid(sDate, 4, 2))
        iDay = CInt(Right(sDate, 2))
        sMon = LCase(Mid(sDate, 6, 3))
        If MonthDict.Exists(sMon) Then
            iMon = MonthDict(sMon)
            dTmp = MostRecent(DateSerial(Year(t.BookDate), iMon, iDay))
            If dTmp <> NODATE Then
                dTmp = dTmp + TimeSerial(iHour, iMin, 0)
                t.TxnDateValid = True
                t.TxnDate = dTmp
            End If
        End If
    End If
End Function
