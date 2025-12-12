' MT2OFX Input Processing Script Argenta (Belgium) CSV format

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/ArgentaBE-CSV.vbs 12    20/02/11 12:08 Colin $"

Const ScriptName = "ArgentaBE-CSV"
Const FormatName = "Argenta (Belgium) CSV Format"
Const ParseErrorMessage = "Cannot parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

Const DebugRecognition = False	' enables debug code in recognition
Const BankCode = "ARSPBE22"
Const CSVSeparator = ";"
Dim NumFieldsExpected
Const DateSequence = "DMY"	' must be DMY, MDY, or YMD
Const DateSeparator = "/-"	' can be empty for dates in e.g. "yyyymmdd" format
Const CurrencyCode = "EUR"	' default if not specified in file
Dim AccountNum				' default if not specified in file
Const NoAvailableBalance = True		' True if file does not contain "Available Balance" information
Dim	bTxnDateYMD				' True if trans date in description is YMD (otherwise DMY)

Const SkipHeaderLines = 1	' number of lines to skip before the transaction data
Const ColumnHeadersPresent = True	' are the column headers in the file?
Const DecimalSeparator = ","	' as used in amounts
Const MemoChunkLength = 0	' if memo field consists of fixed length chunks
Const TxnDatePattern = ".* (\d\d)-(\d\d)-(\d\d)(  \d\d:\d\d)?"	' pattern to find transaction date in the memo
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

' Verrichtingsnummer;Verrichtingsdatum;Omschrijving;Bedrag van de beweg.;Valuta;Valuta datum;Rekening tegenpartij;Naam tegenpartij;Mededeling;Mededeling (vervolg);Verrichtingsreferentie

' Declare fields in the order they appear in the file as an array of arrays. The inner arrays
' contain a field ID from the list above followed by the exact column header.
Dim aFields, aFieldsOld, aFieldsNew, aFieldsNew2
aFieldsOld = Array( _
	Array(fldSkip, "Verrichtingsnummer"), _
	Array(fldBookDate, "Verrichtingsdatum"), _
	Array(fldSkip, "Omschrijving"), _
	Array(fldAmount, "Bedrag van de beweg."), _
	Array(fldCurrency, "Valuta"), _
	Array(fldValueDate, "Valuta datum"), _
	Array(fldSkip, "Rekening tegenpartij"), _
	Array(fldPayee, "Naam tegenpartij"), _
	Array(fldMemo, "Mededeling"), _
	Array(fldMemo, "Mededeling (vervolg)"), _
	Array(fldFITID, "Verrichtingsreferentie") _
)

' Nr v/d verrichting;Datum v. verrichting;Beschrijving;Bedrag v/d verrichting;Munt;Valutadatum;Rekening tegenpartij;Naam v/d tegenpartij :;Mededeling 1 :;Mededeling 2 :;Ref. v/d verrichting
aFieldsNew = Array( _
	Array(fldSkip, "Nr v/d verrichting"), _
	Array(fldBookDate, "Datum v. verrichting"), _
	Array(fldSkip, "Beschrijving"), _
	Array(fldAmount, "Bedrag v/d verrichting"), _
	Array(fldCurrency, "Munt"), _
	Array(fldValueDate, "Valutadatum"), _
	Array(fldSkip, "Rekening tegenpartij"), _
	Array(fldPayee, "Naam v/d tegenpartij :"), _
	Array(fldMemo, "Mededeling 1 :"), _
	Array(fldMemo, "Mededeling 2 :"), _
	Array(fldFITID, "Ref. v/d verrichting") _
)

' Valutadatum;Ref. v/d verrichting;Beschrijving;Bedrag v/d verrichting;Munt;Datum v. verrichting;Rekening tegenpartij;Naam v/d tegenpartij :;Mededeling 1 :;Mededeling 2 :
aFieldsNew2 = Array( _
	Array(fldValueDate, "Valutadatum"), _
	Array(fldFITID, "Ref. v/d verrichting"), _
	Array(fldSkip, "Beschrijving"), _
	Array(fldAmount, "Bedrag v/d verrichting"), _
	Array(fldCurrency, "Munt"), _
	Array(fldBookDate, "Datum v. verrichting"), _
	Array(fldSkip, "Rekening tegenpartij"), _
	Array(fldPayee, "Naam v/d tegenpartij :"), _
	Array(fldMemo, "Mededeling 1 :"), _
	Array(fldMemo, "Mededeling 2 :") _
)

aFields = aFieldsNew


' have we been through the recognition process?
Dim bRecognitionDone: bRecognitionDone = False

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
	Array("TxnDateYMD", "Trans.datum is JMD", _
		"Aanvinken als de transactiedatum het formaat J-M-D heeft. Anders wordt D-M-J aangenomen.", _
		ptBoolean) _
	)

Sub Initialise()
    LogProgress ScriptVersion, "Initialise"
	If Not CheckVersion() Then
		Abort
	End If
' fill field lookup dictionary
' NB: only the last occurrence is remembered!
	Dim i
' ArgentaBE: do this based on the new fields for now. If we have to deal with an old-style file it must be
' patched up after recgnition
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

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
	Dim vFields
	Dim sLine
	Dim i, iTmp
	Dim bTmp
	Dim sField
	Dim sPat
	RecogniseTextFile = False
	For i=1 To SkipHeaderLines
		If AtEOF() Then
			Exit Function
		End If
		sLine = ReadLine()
' ArgentaBE: first line contains account number
		If i=1 Then
'Rekeningnummer :;979-9761356-57;Giro +;;;;;;;;
			If StartsWith(sLine, "Rekeningnummer :;") Then
				AccountNum = Mid(sLine, 18)
			ElseIf StartsWith(sLine, "Nr v/d rekening :;") Then
'Nr v/d rekening :;096-9280001-32;Golden;;;;;;;;
				AccountNum = Mid(sLine, 19)
			Else
				If DebugRecognition Then
					MsgBox "Missing account number in first line"
				End If
				Exit Function
			End If
         iTmp = InStr(AccountNum, ";")
         If iTmp > 0 Then
            AccountNum = Left(AccountNum, iTmp-1)
         End If
         AccountNum = Replace(AccountNum, " ", "")
		End If
	Next
	If AtEOF() Then
		Exit Function
	End If
	sLine = ReadLine()
	vFields = ParseLineDelimited(sLine, CSVSeparator)
	If TypeName(vFields) <> "Variant()" Then
		If DebugRecognition Then
			MsgBox "not var array"
		End If
		Exit Function
	End If
' ArgentaBE: detect old vs. new format. We assumed new format before. If it is old format we must rebuild
' the field map. NB: old and new differ in the first column heading!
	If vFields(1) = aFieldsOld(0)(1) Then
		LogProgress ScriptName, "Old Format File Recognised"
		If DebugRecognition Then
			MsgBox "ArgentaBE: old format"
		End If
		aFields = aFieldsOld
		FieldDict.RemoveAll
		For i=0 To UBound(aFields)
			FieldDict(aFields(i)(0)) = i+1
		Next
	End If
	If vFields(1) = aFieldsNew2(0)(1) Then
		LogProgress ScriptName, "New2 Format File Recognised"
		If DebugRecognition Then
			MsgBox "ArgentaBE: new2 format"
		End If
		aFields = aFieldsNew2
		FieldDict.RemoveAll
		For i=0 To UBound(aFields)
			FieldDict(aFields(i)(0)) = i+1
		Next
	End If
   NumFieldsExpected = UBound(aFields)+1
	If UBound(vFields) <> NumFieldsExpected Then
		If DebugRecognition Then
			MsgBox "wrong number of fields - got " & UBound(vFields) & ", expected " _
			& NumFieldsExpected & " - " & sLine
		End If
		Exit function
	End If
	If ColumnHeadersPresent Then
		For i=1 To NumFieldsExpected
			If vFields(i) <> aFields(i-1)(1) Then
				If DebugRecognition Then
					MsgBox "field " & CStr(i) & " " & aFields(i-1)(1) & " instead of " & vFields(i)
				End If
				Exit function
			End if
		Next
	Else
' pattern-match the first row
		For i=1 To UBound(aFields)-1
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
						sPat = "[ 0-9,]*(\.[0-9]*)?"
					Else
						sPat = "[ 0-9\.]*(,[0-9]*)?"
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
				Exit Function
			End If
		Next
	End If
	LogProgress ScriptName, "File Recognised"
	bRecognitionDone = True
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
   Dim da: Set da = New DateAccumulator

	LoadTextFile = False
	
	' if we haven't been through recognition, just check it now	
	If Not bRecognitionDone Then
		If Not RecogniseTextFile() Then
			Abort
			Exit Function
		End If
		Rewind
	End If

	bTxnDateYMD = GetProperty("TxnDateYMD")
	sAcct = ""
	For i=1 To SkipHeaderLines
		sLine = ReadLine()
	Next
	If ColumnHeadersPresent Then
		sLine = ReadLine()
	End if
	Do While Not AtEOF()
		sLine = ReadLine()
		If Len(sLine) > 0 And Left(sLine,1) <> CSVSeparator Then
' Argenta: place names can legitimately start with a single quote and these are not wrapped in double quotes
			sLine = FixQuotes(sLine)
			vFields = ParseLineDelimited(sLine, CSVSeparator)
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			If UBound(vFields) <> NumFieldsExpected Then
				Message True, True, "Wrong number of fields - " & CStr(UBound(vFields)+1) & " - '" & sLine & "'", ScriptName
				Abort
				Exit function
			End If
	' set up new transaction, and start a new statement if the account # changes
			If FieldDict.Exists(fldAccountNum) Then
				If sAcct <> vFields(FieldDict(fldAccountNum)) Then
					Set Stmt = NewStatement()
		' this initialisation should be in the class constructor!! (fixed in 3.3.5)
					Stmt.OpeningBalance.BalDate = NODATE
					Stmt.OpeningBalance.Ccy = CurrencyCode
					Stmt.AvailableBalance.BalDate = NODATE
					If Not NoAvailableBalance Then Stmt.AvailableBalance.Ccy = CurrencyCode
					Stmt.ClosingBalance.BalDate = NODATE				
					Stmt.ClosingBalance.Ccy = CurrencyCode
					iSeq = 0
					Stmt.BankName = BankCode
				End If
			Else
				If IsEmpty(Stmt) Then
					Set Stmt = NewStatement()
		' this initialisation should be in the class constructor!! (fixed in 3.3.5)
					Stmt.OpeningBalance.BalDate = NODATE
					Stmt.OpeningBalance.Ccy = CurrencyCode
					Stmt.AvailableBalance.BalDate = NODATE
					If Not NoAvailableBalance Then Stmt.AvailableBalance.Ccy = CurrencyCode
					Stmt.ClosingBalance.BalDate = NODATE				
					Stmt.ClosingBalance.Ccy = CurrencyCode
					iSeq = 0
					Stmt.BankName = BankCode
					Stmt.Acct = AccountNum
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
					If Not NoAvailableBalance Then Stmt.AvailableBalance.Ccy = sField
				case fldClosingBal
					Stmt.ClosingBalance.Amt = ParseNumber(sField, DecimalSeparator)
				case fldAvailBal
					Stmt.AvailableBalance.Amt = ParseNumber(sField, DecimalSeparator)
				case fldBookDate
					Txn.BookDate = ParseDate(sField)
               da.Process Txn.BookDate
				case fldValueDate
					Txn.ValueDate = ParseDate(sField)
               da.Process Txn.ValueDate
				Case fldTransactionDate
					Txn.TxnDate = ParseDate(sField)
					Txn.TxnDateValid = (Txn.TxnDate <> NODATE)
               da.Process Txn.TxnDate
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
                If UBound(vDateBits) >= 3 Then
                    If bTxnDateYMD Then
                        Txn.TxnDate = DateSerial(CInt(vDateBits(1))+2000, CInt(vDateBits(2)), CInt(vDateBits(3)))
                    Else
                        Txn.TxnDate = DateSerial(CInt(vDateBits(3))+2000, CInt(vDateBits(2)), CInt(vDateBits(1)))
                    End If
                    Txn.TxnDateValid = True
                End If
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

' keep tabs on the statement/balance Date
			If Stmt.ClosingBalance.BalDate = NODATE Then
				Stmt.ClosingBalance.BalDate = da.MaxDate
			End If
			If Stmt.OpeningBalance.BalDate = NODATE Then
				Stmt.OpeningBalance.BalDate = da.MinDate
			End If

' Special For Argenta BE
			sTmp = TransType(Trim(vFields(3)), Txn.Amt)
			If Len(sTmp) > 0 Then
				Txn.TxnType = sTmp
			End If
			If Trim(vFields(3)) = "Rentenota" Then
				Txn.Memo = "Rentenota"
			End If
		End If
	Loop
	LoadTextFile = True
End Function

Function TransType(sTxt, dAmt)
	Dim sTmp
	Select Case sTxt
	Case "Overschrijving via internet"
		If dAmt > 0 Then
			sTmp = "DEP"
		Else
			sTmp = "PAYMENT"
		End If
	Case "Betaling Bancontact"
		sTmp = "POS"
	Case "Opname Bancontact"
		sTmp = "ATM"
	Case "Debet ten gunste van BCC"
		sTmp = "PAYMENT"
	Case "Overschrijving in uw voordeel"
		sTmp = "CREDIT"
	Case "Uw doorlopende opdracht"
		sTmp = "REPEATPMT"
	Case "Gedomicilieerde factuur"
		sTmp = "DIRECTDEBIT"
	Case "opladen proton"
		sTmp = "PAYMENT"
	Case "Rentenota"
		sTmp = "INT"
	Case "Aankoop brandstof Bancontact"
		sTmp = "POS"
	Case "Betaling België"
		sTmp = "POS"
	Case Else
		sTmp = ""
	End Select
	TransType = sTmp
End Function

' Argenta: place names can legitimately start with a single quote and these are not wrapped in double quotes

Const FixQuotesPattern = "(.*;)('[^;]*)(;.*)"
Const FixQuotesReplace = "$1""$2""$3"

Function FixQuotes(sLine)
	Dim r
	Set r=New RegExp
	r.Global = True
	r.Pattern = FixQuotesPattern
	r.IgnoreCase = True
	FixQuotes = r.Replace(sLine, FixQuotesReplace)
End Function

' =======================================
' OFX Transaction Types
' Type        Description
' CREDIT      Generic credit
' DEBIT       Generic debit
' INT         Interest earned or paid
'             Note: Depends on signage of amount
' DIV         Dividend
' FEE         FI fee
' SRVCHG      Service charge
' DEP         deposit
' ATM         ATM debit or credit
'             Note: Depends on signage of amount
' POS         Point of sale debit or credit
'             Note: Depends on signage of amount
' XFER        Transfer
' CHECK       Check
' PAYMENT     Electronic payment
' CASH        Cash withdrawal
' DIRECTDEP   Direct deposit
' DIRECTDEBIT Merchant initiated debit
' REPEATPMT   Repeating payment/standing order
' OTHER       Other
' =======================================
