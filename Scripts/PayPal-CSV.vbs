' MT2OFX Input Processing Script PayPal CSV format

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/PayPal-CSV.vbs 2     30/01/11 13:31 Colin $"

Dim Params: Set Params = New MT2OFXScript

With Params
	.MinimumProgramVersion = "3.5.37"
	.DebugRecognition = true	' enables debug code in recognition
	.ScriptName = "PayPal-CSV"
	.FormatName = "PayPal CSV"
	.ParseErrorMessage = "Cannot parse line."
	.ParseErrorTitle = .ScriptName
	.BankCode = "PayPal"
	.AccountNum = ""		' default if not specified in file
	.BranchCode = ""		' default if not specified in file
	.QuickenBankID = ""		' copied to INTU.BID if present
	.CurrencyCode = "EUR"	' default if not specified in file
	.ColumnHeadersPresent = True	' are the column headers in the file?
	.SkipHeaderLines = 0	' number of lines to skip before the transaction data
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
	.CSVSeparator = ","
	.DecimalSeparator = ","	' as used in amounts
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
	.MonthNames = Empty
' PayPal allows user selection of certain fields. The .Fields array is built dynamically based on the
' fields actually present in the file.
	.Fields = Array()
' min/max fields expected: default to size of Fields array. can be overridden here if required
'	.MinFieldsExpected = 1
'	.MaxFieldsExpected = 1
	.Properties = Array( _
		Array("AcctNum", "Account number", _
			"The account number for " & ScriptName, _
			ptString,,"=CheckAccount", "Please enter a valid account number.") _
		)
	Set .TransactionCallback = GetRef("TransactionCallback")
'	Set .IsValidTxnLine = GetRef("IsValidTxnLine")
	Set .PreParseCallback = GetRef("PreParseCallback")
'	Set .HeaderCallback = GetRef("HeaderCallback")
	Set .StatementCallback = GetRef("StatementCallback")
'	Set .CustomDateCallback = GetRef("CustomDateCallback")
'	Set .CustomAmountCallback = GetRef("CustomAmountCallback")
'	Set .ReadLineCallback = GetRef("ReadLineCallback")
End With

Dim PayPalFieldList, ppFee, ppGross, ppTimeZone, ppStatus, ppCurrency, ppType, ppCountry

' NB: City, State, PostalCode, Country and Phone are not used in QIF

' PayPalList allows lookup by header and returns a field info Array
Dim PayPalList: Set PayPalList = CreateObject("Scripting.Dictionary")
' PayPalColumns allows lookup by header and returns a vFields index
Dim PayPalColumns: Set PayPalColumns = CreateObject("Scripting.Dictionary")

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
' PayPal: Dynamically fill Fields array based on first line
Function RecogniseTextFile()
    Session.InputFile.Codepage = CP_UTF8
    Dim sLine: sLine=ReadLine()
' check for German etc
    SetupLanguage sLine
' maybe switch to CP1252?
    If ppFee = "Gebühr" And Not InStr(sLine, "ä") Then
        Session.InputFile.Codepage = CP_ACP
        Session.InputFile.Rewind
        sLine = ReadLine()
    End If

' handle tab-separated files as well!
    If InStr(sLine, vbTab) > 0 Then
        Params.CSVSeparator = vbTab
    End If

    Dim vFields: vFields=ParseLineDelimited(sLine, Params.CSVSeparator)
    Dim i, v, sTmp
    Dim nFields: nFields = UBound(vFields) - LBound(vFields) + 1

	RecogniseTextFile = False

	If nFields < 5 Then
		Exit Function
	End If
	
    Dim NewFields(): ReDim NewFields(nFields-1)

    For i=LBound(vFields) To UBound(vFields)
        sTmp = Trim(vFields(i))
        If PayPalList.Exists(sTmp) Then
            NewFields(i-1) = PayPalList(sTmp)
            PayPalColumns(sTmp) = i
        Else
            If Params.DebugRecognition Then
                MsgBox "Unknown PayPal field: " & sTmp
            End If
            Exit Function
        End If
    Next
    Params.Fields = NewFields
    Params.MinFieldsExpected = nFields - 1	' sometimes a final empty field gets missed
'	Params.MaxFieldsExpected = nFields
    Rewind
	
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
    Dim sTZ, lBias, bDoSplits, dFee, dGross, oSplit
'MsgBox "In transaction callback: " & t.Memo
	t.BookDate = t.TxnDate
' all dates in the English paypal file are (US) Pacific Time: either PST or PDT, according to the column "Time Zone"
' here we correct the time, initially back to GMT, then from there to local time
    If PayPalColumns.Exists(ppTimeZone) Then
        sTZ = vFields(PayPalColumns(ppTimeZone))
        If sTZ = "PDT" Then
            lBias = 8 * 60
        ElseIf sTZ = "PST" Then
            lBias = 7 * 60
        ElseIf sTz = "MEZ" Then
            lBias = -1 * 60
        ElseIf sTz = "MESZ" Then
            lBias = -2 * 60
        Else
            lBias = 0
        End If
        lBias = lBias - LocalTZOffset()
        If t.BookDate <> NODATE Then t.BookDate = DateAdd("n", lBias, t.BookDate)
        If t.TxnDate <> NODATE Then t.TxnDate = DateAdd("n", lBias, t.TxnDate)
        If t.ValueDate <> NODATE Then t.ValueDate = DateAdd("n", lBias, t.ValueDate)
    End If
    If PayPalColumns.Exists(ppFee) Or PayPalColumns.Exists(ppGross) Then
        If PayPalColumns.Exists(ppFee) Then
            dFee = ParseNumber(vFields(PayPalColumns(ppFee)), Params.DecimalSeparator)
            If PayPalColumns.Exists(ppGross) Then
                dGross = ParseNumber(vFields(PayPalColumns(ppGross)), Params.DecimalSeparator)
            Else
                dGross = t.Amt - dFee
            End If
        Else
            dGross = ParseNumber(vFields(PayPalColumns(ppGross)), Params.DecimalSeparator)
            dFee = t.Amt - dGross
        End If
' only make splits if there is a fee, otherwise net=gross anyway...
        If dFee <> 0.0 Then
            Set oSplit = t.Splits.AddNew
            oSplit.Memo = ppFee
            oSplit.Amt = dFee
            Set oSplit = t.Splits.AddNew
            oSplit.Memo = ppGross
            oSplit.Amt = dGross
            If Len(t.Memo) > 0 Then
                t.Memo = t.Memo & Cfg.MemoDelimiter
            End If
            t.Memo = t.Memo & ppGross & " " & FormatNumber(dGross, 2) & " " & ppFee & " " & FormatNumber(dFee, 2)
        End If
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
' "Type" can be one of the following:
'  Add Funds from a Bank Account
'  Authorization
'  Currency Conversion
'  eBay Payment Sent
'  Express Checkout Payment Sent
'  Payment Received
'  Payment Sent
'  Preapproved Payment Sent
'  Refund
'  Shopping Cart Item
'  Shopping Cart Payment Sent
'  Subscription Payment Sent
'  Temporary Hold
'  Update to eCheck Received
'  Update to Reversal
'  Web Accept Payment Received
'  Web Accept Payment Sent
'  Withdraw Funds to a Bank Account

Function PreParseCallback(vFields)
'MsgBox "In preparse callback: " & UBound(vFields) & " fields."
	Dim sTmp
	PreParseCallback = False

	Dim sStatus, sType, sCurr
	If PayPalColumns.Exists(ppStatus) Then
		sStatus = vFields(PayPalColumns(ppStatus))
	End If
	If sStatus <> "Completed" And sStatus <> "Refunded" And sStatus <> "Partially Refunded" And sStatus <> "Cleared" Then
		Exit Function
	End If
	If PayPalColumns.Exists(ppCurrency) Then
		sCurr = vFields(PayPalColumns(ppCurrency))
      SetContext sCurr
	End If
	If PayPalColumns.Exists(ppType) Then
		sType = vFields(PayPalColumns(ppType))
	End If
	Select Case sType
	Case "Add Funds from a Bank Account"
		PreParseCallback = True
	Case "Authorization"
	Case "Currency Conversion"
        PreParseCallback = True
    Case "Donation Sent", "Donation Received"
        PreParseCallback = True
    Case "eBay Payment Sent"
        PreParseCallback = True
    Case "eCheck Received"
        PreParseCallback = True
	Case "Express Checkout Payment Sent"
		PreParseCallback = True
    Case "Order"
	Case "Payment Received"
		PreParseCallback = True
	Case "Payment Sent"
		PreParseCallback = True
	Case "Preapproved Payment Sent"
		PreParseCallback = True
    Case "Recurring Payment Sent", "Recurring Payment Received"
        PreParseCallback = True
	Case "Refund"
		PreParseCallback = True
	Case "Shopping Cart Item"
' this is a "split" - cannot process at the moment
	Case "Shopping Cart Payment Sent"
		PreParseCallback = True
	Case "Subscription Payment Sent"
		PreParseCallback = True
	Case "Temporary Hold"
   Case "Transfer"
		PreParseCallback = True   
	Case "Update to eCheck Received"
   Case "Update to Payment Received"
	Case "Update to Reversal"
	Case "Web Accept Payment Received"
		PreParseCallback = True
	Case "Web Accept Payment Sent"
		PreParseCallback = True
	Case "Withdraw Funds to a Bank Account"
		PreParseCallback = True
    case else
        MsgBox "Unknown trans type: " & sType
	End Select
	If PayPalColumns.Exists(ppCountry) Then
		sTmp = MapCountryToISO3166(vFields(PayPalColumns(ppCountry)))
		If Len(sTmp) > 0 Then vFields(PayPalColumns(ppCountry)) = sTmp
	End If	
End Function
Function FinaliseCallback()
MsgBox "In finalisation callback"
	FinaliseCallback = True
End Function
Function IsValidTxnLine(sLine)
MsgBox "In IsValidTxnLine callback: " & sLine
	IsValidTxnLine = (IsNumeric(Left(sLine, 1)))
End Function

Function StatementCallback(Stmt)
' all dates in the paypal file are (US) Pacific Time: either PST or PDT, according to the column "Time Zone"
' here we correct the time, initially back to GMT, then from there to local time
    Dim t, da
    Set da = New DateAccumulator
    For Each t In Stmt.Txns
        da.Process t.BookDate
    Next
    Stmt.OpeningBalance.BalDate = da.MinDate
    Stmt.ClosingBalance.BalDate = da.MaxDate
' MsgBox "In statement callback"
    Stmt.Acct = "Paypal - " & Stmt.ClosingBalance.Ccy

StatementCallback = True
End Function

Sub SetupLanguage(sLine)
    Dim i, sTmp
    If InStr(sLine, "Zeitzone") Then
        GoGerman
    Else
        GoEnglish
    End If
    ' fill a dictionary to facilitate looking up fields by header text
    For i=LBound(PayPalFieldList) To UBound(PayPalFieldList)
        sTmp = PayPalFieldList(i)(1)
' If i<10 then MsgBox "fld: " & sTmp
        PayPalList(sTmp) = PayPalFieldList(i)
    Next
End Sub

Sub GoEnglish()
    ppFee = "Fee"
    ppGross = "Gross"
    ppTimeZone = "Time Zone"
    ppStatus = "Status"
    ppCurrency = "Currency"
    ppType = "Type"
    ppCountry = "Country"
    PayPalFieldList = Array( _
		Array(fldTransactionDate, "Date"), _
		Array(fldTransactionTime, "Time"), _
		Array(fldSkip, ppTimeZone), _
		Array(fldPayee, "Name"), _
		Array(fldSkip, ppType), _
		Array(fldSkip, ppStatus), _
		Array(fldMemo, "Subject"), _
		Array(fldCurrency, ppCurrency), _
		Array(fldSkip, ppGross), _
		Array(fldSkip, ppFee), _
		Array(fldAmount, "Net"), _
		Array(fldSkip, "Note"), _
		Array(fldSkip, "From Email Address"), _
		Array(fldSkip, "To Email Address"), _
		Array(fldFITID, "Transaction ID"), _
		Array(fldSkip, "Payment Type"), _
		Array(fldSkip, "Counterparty Status"), _
		Array(fldSkip, "Shipping Address"), _
		Array(fldSkip, "Address Status"), _
		Array(fldSkip, "Item Title"), _
		Array(fldSkip, "Item ID"), _
		Array(fldSkip, "Shipping and Handling Amount"), _
		Array(fldSkip, "Insurance Amount"), _
		Array(fldSkip, "Sales Tax"), _
		Array(fldSkip, "Option 1 Name"), _
		Array(fldSkip, "Option 1 Value"), _
		Array(fldSkip, "Option 2 Name"), _
		Array(fldSkip, "Option 2 Value"), _
		Array(fldSkip, "Auction Site"), _
		Array(fldSkip, "Buyer ID"), _
		Array(fldSkip, "Item URL"), _
		Array(fldSkip, "Closing Date"), _
		Array(fldSkip, "Escrow Id"), _
		Array(fldSkip, "Reference Txn ID"), _
		Array(fldSkip, "Invoice Number"), _
		Array(fldSkip, "Invoice Id"), _
		Array(fldSkip, "Subscription Number"), _
		Array(fldSkip, "Custom Number"), _
		Array(fldSkip, "Quantity"), _
		Array(fldSkip, "Receipt ID"), _
		Array(fldClosingBal, "Balance"), _
		Array(fldPayeeAddress1, "Address Line 1"), _
		Array(fldPayeeAddress2, "Address Line 2/District"), _
		Array(fldPayeeAddress2, "Address Line 2/District/Neighborhood"), _
		Array(fldPayeeCity, "Town/City"), _
		Array(fldPayeeState, "State/Province/Region/County/Territory/Prefecture/Republic"), _
		Array(fldPayeeZip, "Zip/Postal Code"), _
		Array(fldPayeeCountry, ppCountry), _
		Array(fldPayeePhone, "Contact Phone Number"), _
		Array(fldSkip, "Balance Impact"), _
		Array(fldEmpty, "") _
	)
End Sub
Sub GoGerman()
    ppFee = "Gebühr"
    ppGross = "Brutto"
    ppTimeZone = "Zeitzone"
    ppStatus = "Status"
    ppCurrency = "Währung"
    ppType = "Art"
    ppCountry = "Land"
    PayPalFieldList = Array( _
		Array(fldTransactionDate, "Datum"), _
		Array(fldTransactionTime, "Zeit"), _
		Array(fldSkip, ppTimeZone), _
		Array(fldPayee, "Name"), _
		Array(fldSkip, ppType), _
		Array(fldSkip, ppStatus), _
		Array(fldCurrency, ppCurrency), _
		Array(fldSkip, ppGross), _
		Array(fldSkip, ppFee), _
		Array(fldAmount, "Netto"), _
		Array(fldSkip, "Von E-Mail-Adresse"), _
		Array(fldSkip, "An E-Mail-Adresse"), _
		Array(fldFITID, "Transaktionscode"), _
		Array(fldSkip, "Status der Gegenpartei"), _
		Array(fldSkip, "Adressstatus"), _
		Array(fldSkip, "Verwendungszweck"), _
		Array(fldSkip, "Artikelnummer"), _
		Array(fldSkip, "Betrag für Versandkosten"), _
		Array(fldSkip, "Versicherungsbetrag"), _
		Array(fldSkip, "Umsatzsteuer"), _
		Array(fldSkip, "Option 1 - Name"), _
		Array(fldSkip, "Option 1 - Wert"), _
		Array(fldSkip, "Option 2 - Name"), _
		Array(fldSkip, "Option 2 - Wert"), _
		Array(fldSkip, "Auktions-Site"), _
		Array(fldSkip, "Käufer-ID"), _
		Array(fldSkip, "Artikel-URL"), _
		Array(fldSkip, "Angebotsende"), _
		Array(fldSkip, "Vorgangs-Nr."), _
		Array(fldSkip, "Txn-Referenzkennung"), _
		Array(fldSkip, "Rechnungs-Nr."), _
		Array(fldSkip, "Rechnungsnummer"), _
		Array(fldSkip, "Individuelle Nummer"), _
		Array(fldSkip, "Bestätigungsnummer"), _
		Array(fldPayeeAddress1, "Adresse"), _
		Array(fldPayeeAddress2, "Zusätzliche Angaben/Adresszusatz"), _
		Array(fldPayeeCity, "Stadt"), _
		Array(fldPayeeState, "Staat/Provinz/Region/Landkreis/Territorium/Präfektur/Republik"), _
		Array(fldPayeeZip, "PLZ"), _
		Array(fldPayeeCountry, ppCountry), _
		Array(fldPayeePhone, "Telefonnummer der Kontaktperson"), _
		Array(fldEmpty, "") _
	)
'Datum	 Zeit	 Zeitzone	 Name	 Art	 Status	 Währung	 Brutto	 Gebühr	 Netto	 Von E-Mail-Adresse	 An E-Mail-Adresse	 Transaktionscode	 Status der Gegenpartei	 Adressstatus	 Verwendungszweck	 Artikelnummer	 Betrag für Versandkosten	 Versicherungsbetrag	 Umsatzsteuer	 Option 1 - Name	 Option 1 - Wert	 Option 2 - Name	 Option 2 - Wert	 Auktions-Site	 Käufer-ID	 Artikel-URL	 Angebotsende	 Vorgangs-Nr.	 Rechnungs-Nr.	 Txn-Referenzkennung	 Rechnungsnummer	 Individuelle Nummer	 Bestätigungsnummer	 Adresse	 Zusätzliche Angaben/Adresszusatz	 Stadt	 Staat/Provinz/Region/Landkreis/Territorium/Präfektur/Republik	 PLZ	 Land	 Telefonnummer der Kontaktperson
End Sub
