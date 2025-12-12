' MT2OFX Input Processing Script Basic CSV format
' NB: This Script Will Not Work Without Customisation!

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/Norma43ES-TXT.vbs 2     30/01/11 13:29 Colin $"

Dim Params: Set Params = New MT2OFXScript

With Params
	.MinimumProgramVersion = "3"
	.DebugRecognition = False	' enables debug code in recognition
	.ScriptName = "Norma43"
	.FormatName = "Norma43 (Spain)"
	.ParseErrorMessage = "Cannot parse line."
	.ParseErrorTitle = .ScriptName
   .CodePage = 850  ' Recommended for norma43 in ASCII
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
	.CSVSeparator = ""
	.DecimalSeparator = ","	' as used in amounts
	.TxnLinePattern = ""
	.DateSequence = "YMD"	' must be DMY, MDY, or YMD
	.DateSeparator = ""	' can be empty for dates in e.g. "yyyymmdd" format
	.OldestLast = False		' True if transactions are in reverse order
	.InvertSign = False	' make credits into debits etc
	.NoAvailableBalance = True		' True if file does not contain "Available Balance" information
	.MemoChunkLength = 0	' if memo field consists of fixed length chunks
	.TxnDatePattern = ".*(\d\d)\.(\d\d)\.(\d\d)\ (\d\d)\.(\d\d)"	' pattern to find transaction date in the memo
	.TxnDateSequence = Array(3,2,1,4,5,0)	' order of the info in the pattern (from 1 to 6): Y,M,D,H,M,S
	.PayeeLocation = 0		' start of payee in memo
	.PayeeLength = 0		' length of payee in memo
	.MonthNames = Empty
	.Fields = Array( _
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
    Set .PreParseCallback = GetRef("PreParseCallback")
'	Set .HeaderCallback = GetRef("HeaderCallback")
'	Set .StatementCallback = GetRef("StatementCallback")
'	Set .CustomDateCallback = GetRef("CustomDateCallback")
    Set .CustomAmountCallback = GetRef("CustomAmountCallback")
'	Set .ReadLineCallback = GetRef("ReadLineCallback")
End With

Dim vFields11, vFields22, vFields23, vFields24, vFields33, vFields88
Dim pat11, pat22, pat23,pat24, pat33, pat88

pat11 = "11(\d{4})(\d{4})(\d{10})(\d{6})(\d{6})(\d{1})(\d{14})(\d{3})(\d{1})(.{26})..."
' 111234123400800123451009011011032000000000506889781ACCOUNT HOLDER NAME          
vFields11 = Array( _
    Array(fldSkip, "Clave de la Entidad"), _
    Array(fldBranch, "Clave de Oficina"), _
    Array(fldAccountNum, "No de cuenta"), _
    Array(fldSkip, "Fecha inicial"), _
    Array(fldSkip, "Fecha final"), _
    Array(fldSkip, "Clave Debe o Haber"), _
    Array(fldSkip, "Importe saldo inicial"), _
    Array(fldCurrency, "Clave de divisa", "\d\d\d"), _
    Array(fldSkip, "Modalidad de información"), _
    Array(fldSkip, "Nombre abreviado") _
)

pat22 = "22....(\d{4})(\d{6})(\d{6})(\d\d)(\d{3})(\d)(\d{14})(\d{10})(\d{12})(.{16})"
'22    0080100901100901040742000000000500000000000000000000000000                
vFields22 = Array( _
    Array(fldSkip, "Clave de Oficina Origen"), _
    Array(fldBookDate, "Fecha operación"), _
    Array(fldValueDate, "Fecha valor"), _
    Array(fldSkip, "Concepto común"), _
    Array(fldSkip, "Concepto proprio"), _
    Array(fldSkip, "Clave Debe o Haber"), _
    Array(fldAmount, "Importe"), _
    Array(fldSkip, "No de documento"), _
    Array(fldMemo, "Referencia 1"), _
    Array(fldMemo, "Referencia 2") _
)

pat23 = "23(\d\d)(.{38})(.{38})"
' 2301SANITAS SA                    RPRL01091000385420                            
vFields23 = Array( _
    Array(fldSkip, "Código Dato"), _
    Array(fldMemo, "Concepto"), _
    Array(fldMemo, "Concepto") _
)

pat24 = "24(\d\d))(\d\d\d)(\d{14}).{59}"
vFields24 = Array( _
    Array(fldSkip, "Código Dato"), _
    Array(fldSkip, "Clave divisa origen del movimiento"), _
    Array(fldSkip, "Importe") _
)

pat33 = "33(\d{4})(\d{4})(\d{10})(\d{5})(\d{14})(\d{5})(\d{14})(\d)(\d{14})(\d{3})...."
' 3312341234008001234500014000000001168410000300000000092286200000000026133978    
vFields33 = Array( _
    Array(fldSkip, "Clave de Entidad"), _
    Array(fldSkip, "Clave de Oficina"), _
    Array(fldSkip, "No de cuenta"), _
    Array(fldSkip, "No apuntes Debe"), _
    Array(fldSkip, "Total importes Debe"), _
    Array(fldSkip, "No apuntes Haber"), _
    Array(fldSkip, "Total importes Haber"), _
    Array(fldSkip, "Código Saldo final"), _
    Array(fldClosingBal, "Saldo final"), _
    Array(fldSkip, "Clave de Divisa") _
)

pat88 = "88(\d{18})(\d{5}).*"
' 8899999999999999999900035                                                      
vFields88 = Array( _
    Array(fldSkip, "Nueves"), _
    Array(fldSkip, "No de registros") _
)

Dim sRec

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
    Params.TxnLinePattern = pat11
    Params.Fields = vFields11
    RecogniseTextFile = DefaultRecogniseTextFile(Params)
    If RecogniseTextFile Then
        LogProgress ScriptName, "File Recognised"
    End If
End Function

Function LoadTextFile()
    Params.TxnLinePattern = pat11
    Params.Fields = vFields11
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
' MsgBox "In transaction callback: " & t.Memo
    If sRec = "22" Then
        If vFields(6) = "1" Then t.Amt = -t.Amt
    ElseIf sRec = "23" Then
        If vFields(1) = "01" Then t.Payee = Trim(Left(vFields(2), 30))
    ElseIf sRec = "33" Then
        If vFields(9) = "1" Then Stmt.ClosingBalance.Amt = -Stmt.ClosingBalance.Amt
    End If
    TransactionCallback = True
End Function
Function CustomDateCallback(sDate)
MsgBox "In custom date callback: " & sDate
	CustomDateCallback = ParseDateEx(sDate, Params.DateSequence, Params.DateSeparator)
End Function
Function CustomAmountCallback(sAmt)
'MsgBox "In custom amount callback: " & sAmt
    CustomAmountCallback = ParseNumber(sAmt, Params.DecimalSeparator) / 100.0
End Function
Function ReadLineCallback(sLine)
MsgBox "In read line callback: " & sLine
	ReadLineCallback = sLine
End Function
' PreParseCallback: returns True or False. True means the line can be processed; False means skip this line.
Function PreParseCallback(vFields)
' MsgBox "In preparse callback: " & UBound(vFields) & " fields."
    If sRec = "11" Then
        Params.BankCode = vFields(2) & "ES"  ' quick hack - the bank code is a national 4-digit code
    End If
    PreParseCallback = True
End Function
Sub StatementCallback(Stmt)
MsgBox "In statement callback"
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
'MsgBox "In IsValidTxnLine callback: " & sLine
    sRec = Left(sLine, 2)
    Select Case sRec
    Case "11"
        Params.TxnLinePattern = pat11
        Params.Fields = vFields11
        IsValidTxnLine = txnlineNOTRANSACTION
    Case "22"
        Params.TxnLinePattern = pat22
        Params.Fields = vFields22
        IsValidTxnLine = txnlineNORMAL
    Case "23"
        Params.TxnLinePattern = pat23
        Params.Fields = vFields23
        IsValidTxnLine = txnlineCONTINUE
    Case "24"
        Params.TxnLinePattern = pat24
        Params.Fields = vFields24
        IsValidTxnLine = txnlineCONTINUE
    Case "33"
        Params.TxnLinePattern = pat33
        Params.Fields = vFields33
        IsValidTxnLine = txnlineNOTRANSACTION
    Case "88"
        Params.TxnLinePattern = pat88
        Params.Fields = vFields88
        IsValidTxnLine = txnlineNOTRANSACTION
    Case Else
        IsValidTxnLine = False
    End Select
End Function
