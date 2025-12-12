' MT2OFX Input Processing Script CODA 2.2 format (Belgium)
Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/CodaBE.vbs 3     24/11/09 22:04 Colin $"

' Reusable pattern for dates: DDMMYY
Const csDatePat = "\d{6}"

Dim Pattern0, Pattern1, Pattern21, Pattern22, Pattern23, Pattern31, Pattern32, Pattern33, Pattern8, Pattern9
Dim Fields0, Fields1, Fields21, Fields22, Fields23, Fields31, Fields32, Fields33, Fields8, Fields9
' record type 0: file header
Pattern0 = "00000" & csDatePat & "\d{3}05. {7}.{10}.{26}(.{11})\d{11} \d{5}.{16}.{16} {7}2"
' pattern is specified here for the BIC because this line is used for recognition
Fields0 = Array( _
		Array(fldUser1, "BIC rekeninghoudende bank", "[A-Z][A-Z][A-Z][A-Z]BE..") _
	)

'12009BE44091900000145                  EUR1000061234246690080108STADSBESTUUR              ZICHTREKENING/IN AGENTSCH.         007
' record type 1: previous balance (= opening balance)
Pattern1 = "1\d\d{3}(.{37})(\d{16})(" & csDatePat & ").{26}.{35}\d{3}"
Fields1 = Array( _
		Array(fldUser2, "Rekeningnummer"), _
		Array(fldSkip, "Oud saldo"), _
		Array(fldSkip, "Datum oud saldo") _
    )

' record type 2.1: transaction
Pattern21 = "21\d{4}\d{4}.{21}(\d{16})(" & csDatePat & ")\d{8}\d(.{53})(" & csDatePat & ")\d{3}\d\d \d"
Fields21 = Array( _
    Array(fldAmount, "Bedrag"), _
    Array(fldValueDate, "Valutadatum"), _
    Array(fldMemo, "Mededeling"), _
    Array(fldBookDate, "Boekingsdatum") _
    )

' record type 2.2: transaction part 2
'220002001300000000                                                                                                           0 0
'2200030000                                                                                                   00843246306A16481 0
Pattern22 = "22\d{4}\d{4}(.{53}).{35}(.{11}).{8}.{4}.{4}\d \d"
Fields22 = Array ( _
    Array(fldMemo, "Mededeling"), _
    Array(fldPayeeAcctBank, "BIC tegenpartij") _
    )

' record type 2.3: transaction part 3
Pattern23 = "23\d{4}\d{4}(.{37})(.{35})(.{43})0 \d"
Fields23 = Array( _
    Array(fldUser3, "Rekeningnr tegenpartij"), _
    Array(fldPayee, "Naam Tegenpartij"), _
    Array(fldMemo, "Mededeling") _
    )
    
' record type 3.1: transaction info part 1
Pattern31 = "31\d{4}\d{4}.{21}\d{8}\d(.{73}) {12}\d \d"
Fields31 = Array( _
    Array(fldMemo, "Mededeling") _
    )

' record type 3.2: transaction info part 2
Pattern32 = "32\d{4}\d{4}(.{105}) {10}\d \d"
Fields32 = Array( _
    Array(fldMemo, "Mededeling") _
    )

' record type 3.3: transaction info part 3
Pattern33 = "33\d{4}\d{4}(.{90}) {25}0 \d"
Fields33 = Array( _
    Array(fldMemo, "Mededeling") _
    )

' record type 4: free text info
' ignored

'8009BE44091900000145                  EUR1000061239947690090108000000000000000 00000000000000000000000000000001000061239947690 0
' record type 8: closing balance
Pattern8 = "8\d{3}.{37}(\d{16})(" & csDatePat & ").{64}\d"
Fields8 = Array( _
    Array(fldClosingBal, "Nieuw saldo"), _
    Array(fldBalanceDate, "Datum nieuw saldo") _
    )

' record type 9: file trailer
' ignored

Dim Params: Set Params = New MT2OFXScript

With Params
	.MinimumProgramVersion = "3"
	.DebugRecognition = False	' enables debug code in recognition
	.ScriptName = "Coda2.2"
	.FormatName = "Coda 2.2 (Belgium)"
	.ParseErrorMessage = "Cannot parse line."
	.ParseErrorTitle = .ScriptName
	.BankCode = ""
	.AccountNum = ""		' default if not specified in file
	.BranchCode = ""		' default if not specified in file
	.AccountType = "CHECKING"	' can be CHECKING or CREDITCARD
	.QuickenBankID = ""		' copied to INTU.BID if present
	.CurrencyCode = "EUR"	' default if not specified in file
	.ColumnHeadersPresent = False	' are the column headers in the file?
	.SkipHeaderLines = 0	' number of lines to skip before the transaction data
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
	.CSVSeparator = ","
	.DecimalSeparator = ","	' as used in amounts
	.TxnLinePattern = Pattern0
	.DateSequence = "DMY"	' must be DMY, MDY, or YMD
	.DateSeparator = ""	' can be empty for dates in e.g. "yyyymmdd" format
	.OldestLast = False		' True if transactions are in reverse order
	.InvertSign = False	' make credits into debits etc
	.NoAvailableBalance = True		' True if file does not contain "Available Balance" information
	.MemoChunkLength = 0	' if memo field consists of fixed length chunks
	.TxnDatePattern = ""	' pattern to find transaction date in the memo
	.TxnDateSequence = Array(3,2,1,4,5,0)	' order of the info in the pattern (from 1 to 6): Y,M,D,H,M,S
	.PayeeLocation = 0		' start of payee in memo
	.PayeeLength = 0		' length of payee in memo
	.MonthNames = Empty
	.Fields = Fields0
' min/max fields expected: default to size of Fields array. can be overridden here if required
'	.MinFieldsExpected = 1
'	.MaxFieldsExpected = 1
'	.Properties = Array( _
'		Array("AcctNum", "Account number", _
'			"The account number for " & Params.FormatName, _
'			ptString,,"=CheckAccount", "Please enter a valid account number.") _
'		)
'	Set .TransactionCallback = GetRef("TransactionCallback")
	Set .IsValidTxnLine = GetRef("IsValidTxnLine")
'	Set .PreParseCallback = GetRef("PreParseCallback")
'	Set .HeaderCallback = GetRef("HeaderCallback")
	Set .StatementCallback = GetRef("StatementCallback")
'	Set .CustomDateCallback = GetRef("CustomDateCallback")
	Set .CustomAmountCallback = GetRef("CustomAmountCallback")
'	Set .ReadLineCallback = GetRef("ReadLineCallback")
    Set .User1Callback = GetRef("GetBIC")
    Set .User2Callback = GetRef("GetAccount")
    Set .User3Callback = GetRef("GetPayeeAccount")
End With

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
MsgBox "In header callback: " & sLine
	HeaderCallback = True
End Function
Function TransactionCallback(t, vFields)
MsgBox "In transaction callback: " & t.Memo
	TransactionCallback = True
End Function
Function CustomDateCallback(sDate)
MsgBox "In custom date callback: " & sDate
	CustomDateCallback = ParseDateEx(sDate, Params.DateSequence, Params.DateSeparator)
End Function
Function CustomAmountCallback(sAmt)
    If Len(sAmt) <> 16 Then
        MsgBox "In custom amount callback: " & sAmt
    End If
    Dim iSign
    If Left(sAmt, 1) = "1" Then
        iSign = -1.0
    Else
        iSign = 1.0
    End If
    Dim dAmt
    dAmt = CDbl(Mid(sAmt, 2, 12))
    dAmt = iSign * (dAmt + (CDbl(Right(sAmt, 3)) / 1000.0))
'MsgBox "Amount: " & sAmt & " -> " & dAmt
    CustomAmountCallback = dAmt
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
    Stmt.BankName = Params.BankCode
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
Dim sLast: sLast = ""
Dim sAcctNumType
Function IsValidTxnLine(sLine)
'MsgBox "In IsValidTxnLine callback: " & sLine
    Dim sSeq: sSeq = Mid(sLine, 3, 4)
    Select Case Left(sLine, 1)
    Case "0"   ' file header
        Params.Fields = Fields0
        Params.TxnLinePattern = Pattern0
'msgbox "header line: " & sLine
        Params.BankCode = Mid(sLine, 61, 11)
        IsValidTxnLine = txnlineSKIP
    Case "1"   ' previous balance
        Params.Fields = Fields1
        Params.TxnLinePattern = Pattern1
        sAcctNumType = Mid(sLine, 2, 1)
        IsValidTxnLine = txnlineNOTRANSACTIONNEWSTATEMENT
    Case "2"   ' transaction
        Select Case Mid(sLine, 2, 1)
        Case "1"
            Params.Fields = Fields21
            Params.TxnLinePattern = Pattern21
            IsValidTxnLine = txnlineNORMAL
            If sSeq = sLast Then
                IsValidTxnLine = txnlineSKIP
            End If
            sLast = sSeq
        Case "2"
            Params.Fields = Fields22
            Params.TxnLinePattern = Pattern22
            IsValidTxnLine = txnlineCONTINUE
        Case "3"
            Params.Fields = Fields23
            Params.TxnLinePattern = Pattern23
            IsValidTxnLine = txnlineCONTINUE
        Case Else
            IsValidTxnLine = txnlineSKIP
        End Select
    Case "3"   ' extra information
        Select Case Mid(sLine, 2, 1)
        Case "1"
            Params.Fields = Fields31
            Params.TxnLinePattern = Pattern31
            IsValidTxnLine = txnlineCONTINUE
        Case "2"
            Params.Fields = Fields32
            Params.TxnLinePattern = Pattern32
            IsValidTxnLine = txnlineCONTINUE
        Case "3"
            Params.Fields = Fields33
            Params.TxnLinePattern = Pattern33
            IsValidTxnLine = txnlineCONTINUE
        Case Else
            IsValidTxnLine = txnlineSKIP
        End Select
    Case "4"   ' optional information
        IsValidTxnLine = txnlineSKIP
    Case "8"   ' new balance
        Params.Fields = Fields8
        Params.TxnLinePattern = Pattern8
        IsValidTxnLine = txnlineNOTRANSACTION
    Case "9"   ' file trailer
        IsValidTxnLine = txnlineSKIP
    Case Else
        IsValidTxnLine = txnlineSKIP
    End Select
End Function

' callback for User1 = Bank Code
Function GetBIC(s, t, sField)
'msgbox "BIC: " & sField
' this will be in the first line of the file so there is no statement yet
    Params.BankCode = sField
End Function

' callback for User2 - account number/currency - contained in the second line of the file!
Function GetAccount(s, t, sField)
    Dim sTmp, sCcy
    Select Case sAcctNumType
    Case "0"   ' Belgian domestic
        sTmp = Trim(Left(sField, 12))
        sCcy = Mid(sField, 14, 3)
    Case "1"   ' International
        sTmp = Trim(Left(sField, 34))
        sCcy = Right(sField, 3)
    Case "2"   ' Belgian IBAN
        sTmp = Trim(Left(sField, 31))
        sCcy = Right(sField, 3)
    Case "3"   ' International IBAN
        sTmp = Trim(Left(sField, 34))
        sCcy = Right(sField, 3)
    End Select
    Params.AccountNum = sTmp
'    s.Acct = sTmp
    If Len(Trim(sCcy)) > 0 Then
        s.OpeningBalance.Ccy = sCcy
        s.ClosingBalance.Ccy = sCcy
    End If
End Function

' callback for User3 - account number of payee
Function GetPayeeAccount(s, t, sField)
    Dim sTmp
    sTmp = Trim(sField)
    If Len(sTmp) > 0 Then
        t.Payee.Acct = sTmp
        If Len(t.Payee.BankName) = 0 Then t.Payee.BankName = "Bank"
    End If
End Function
