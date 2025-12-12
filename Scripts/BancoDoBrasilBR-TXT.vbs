' MT2OFX Input Processing Script Banco do Brasil TXT format

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/BancoDoBrasilBR-TXT.vbs 2     13/02/11 20:18 Colin $"

Dim Params: Set Params = New MT2OFXScript

With Params
	.MinimumProgramVersion = "3"
	.DebugRecognition = False ' enables debug code in recognition
	.ScriptName = "BancoDoBrasilBR-TXT"
	.FormatName = "Banco do Brasil"
	.ParseErrorMessage = "Cannot parse line."
	.ParseErrorTitle = .ScriptName
   .CodePage = 1252  ' Windows English / Western Europe
	.BankCode = "BRASBRRJ"
	.AccountNum = ""		' default if not specified in file
	.BranchCode = ""		' default if not specified in file
	.AccountType = "CREDITCARD"	' can be CHECKING or CREDITCARD
	.QuickenBankID = ""		' copied to INTU.BID if present
	.CurrencyCode = "BRL"	' default if not specified in file
	.ColumnHeadersPresent = True	' are the column headers in the file?
	.SkipHeaderLines = 15	' number of lines to skip before the transaction data
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
	.CSVSeparator = ""
	.DecimalSeparator = ","	' as used in amounts
	.TxnLinePattern = "(.{8}) (.{38}) (.{5}) (.{14}) (.{11})"
' 10/01/11 PGTO DEBITO CONTA 4728 00              BR         -2.989,70        0,00
' Data     Transações                             País        Valor R$   Valor US$
' 1        10                                     49         60             75
	.DateSequence = "DMY"	' must be DMY, MDY, or YMD
	.DateSeparator = "-/. "	' can be empty for dates in e.g. "yyyymmdd" format
	.OldestLast = True		' True if transactions are in reverse order
	.InvertSign = True	' make credits into debits etc
	.NoAvailableBalance = True		' True if file does not contain "Available Balance" information
	.MemoChunkLength = 0	' if memo field consists of fixed length chunks
	.TxnDatePattern = ".*(\d\d)\.(\d\d)\.(\d\d)\ (\d\d)\.(\d\d)"	' pattern to find transaction date in the memo
	.TxnDateSequence = Array(3,2,1,4,5,0)	' order of the info in the pattern (from 1 to 6): Y,M,D,H,M,S
	.PayeeLocation = 0		' start of payee in memo
	.PayeeLength = 0		' length of payee in memo
	.MonthNames = Empty
	.Fields = Array( _
		Array(fldBookDate, "Data"), _
		Array(fldMemo, "Transações"), _
		Array(fldSkip, "País"), _
		Array(fldAmount, "Valor R$"), _
		Array(fldSkip, "Valor US$") _
	)
' min/max fields expected: default to size of Fields array. can be overridden here if required
'	.MinFieldsExpected = 1
'	.MaxFieldsExpected = 1
	.Properties = Array( _
		Array("AcctNum", "Account number", _
			"The account number for " & Params.FormatName, _
			ptString,,"=CheckAccount", "Please enter a valid account number.") _
		)
	Set .TransactionCallback = GetRef("TransactionCallback")
	Set .IsValidTxnLine = GetRef("IsValidTxnLine")
'	Set .PreParseCallback = GetRef("PreParseCallback")
	Set .HeaderCallback = GetRef("HeaderCallback")
	Set .StatementCallback = GetRef("StatementCallback")
'	Set .CustomDateCallback = GetRef("CustomDateCallback")
'	Set .CustomAmountCallback = GetRef("CustomAmountCallback")
'	Set .ReadLineCallback = GetRef("ReadLineCallback")
End With

Dim dUsdToBrl, dObal, dCbal, dtStmt: dUsdToBrl = 1.0

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
' MsgBox "In header callback: " & sLine
' Nr.Cartão       : 5522.****.****.9623
' Vencimento      : 25.10.2010
    If StartsWith(sLine, "Nr.Cartão       : ") Then
        Params.AccountNum = Trim(Mid(sLine, 19))
    ElseIf StartsWith(sLine, "Vencimento      : ") Then
        dtStmt = ParseDateEx(Trim(Mid(sLine, 19)), "DMY", ".")
    End If
    HeaderCallback = True
End Function
Function TransactionCallback(t, vFields)
    Dim dAmt
' MsgBox "In transaction callback: " & t.Memo
    t.Payee = Trim(Left(vFields(2), 22))
'MsgBox "In transaction callback, Payee = " & t.Payee
    If t.Amt = 0.0 Then
        dAmt = -Params.ParseAmount(vFields(5))   ' get dollar amount and reverse sign!
        t.Amt = dAmt
        If dAmt < 0 Then
            t.TxnType = "PAYMENT"
        Else
            t.TxnType = "DEP"
        End If
        t.BookingCode = "USD"                   ' slight temporary misuse of this field!
    End If
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
    Dim t
    For Each t in Stmt.Txns
        If t.BookingCode = "USD" Then
            t.Amt = t.Amt * dUsdToBrl
            t.BookingCode = ""
        End If
    Next
    Stmt.OpeningBalance.Amt = dObal
    Stmt.ClosingBalance.Amt = dCbal
    Stmt.ClosingBalance.BalDate = dtStmt
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
Dim bLastWasTxn: bLastWasTxn = False
Dim sLastHeader: sLastHeader = ""
Function IsValidTxnLine(sLine)
'MsgBox "In IsValidTxnLine callback: " & sLine
    Dim sTmp
    If Params.ParseDate(Left(sLine, 10)) <> NODATE Then
        IsValidTxnLine = txnlineNORMAL
        bLastWasTxn = True
    Else
'   Anterior     Créditos     Débitos          R$       utilizado      Atual - R$ 
'   2.989,70 -   2.989,70 +     568,39 =      568,39 -       0,00 =         568,39
'   Saques      débitos      Créditos      Atual U$    conversão      convertido 
'--------------------------------------------------------------------------------
'       0,00 -       0,00 +       0,00 =        0,00   X      0.0 =           0,00
        sTmp = Trim(sLine)
        If sTmp = "Anterior     Créditos     Débitos          R$       utilizado      Atual - R$" Then
            sLastHeader = sTmp
        ElseIf sTmp = "Saques      débitos      Créditos      Atual U$    conversão      convertido" Then
'msgbox "found header leading to exch rate"
            sLastHeader = sTmp
        ElseIf Not StartsWith(sLine, "--------------") Then
            If StartsWith(sLastHeader, "Anterior") Then
                dObal = -Params.ParseAmount(Left(sLine, 10))
                dCbal = -Params.ParseAmount(Right(sLine, 14))
            ElseIf StartsWith(sLastHeader, "Saques") Then
                dCbal = dCbal + (-Params.ParseAmount(Right(sLine, 14)))
                dUsdToBrl = ParseNumber(Trim(Mid(sLine, 56, 8)), ".")
'msgbox "exchange rate = " & dUsdToBrl
            End If
            sLastHeader = ""
        End If
        If bLastWasTxn Then
            ConcatMemo(Mid(sLine, 10, 38))
        End If
        bLastWasTxn = False
        IsValidTxnLine = txnlineSKIP
    End If
End Function
