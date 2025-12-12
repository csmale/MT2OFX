' MT2OFX Input Processing Script Nationwide current account CSV format

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/Nationwide-CSV.vbs 9     30/01/11 13:27 Colin $"

Const ScriptName = "Nationwide-CSV"
Const FormatName = "Nationwide current account CSV"
Const ParseErrorMessage = "Cannot parse line."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

Const DebugRecognition = False	' enables debug code in recognition
Const BankCode = "NAIAGB21"
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
Const CSVSeparator = ","
Const TxnLinePattern = ""
Const MinFieldsExpected = 5
Const MaxFieldsExpected = 6
Const DateSequence = "DMY"	' must be DMY, MDY, or YMD
Const DateSeparator = "-/. "	' can be empty for dates in e.g. "yyyymmdd" format
Const InvertSign = False	' make credits into debits etc
Const CurrencyCode = "GBP"	' default if not specified in file
Const NoAvailableBalance = True		' True if file does not contain "Available Balance" information
Dim AccountNum: AccountNum = ""		' default if not specified in file
Dim BranchCode: BranchCode = ""		' default if not specified in file
Const SkipHeaderLines = 9	' number of lines to skip before the transaction data
Const ColumnHeadersPresent = True	' are the column headers in the file?
Const DecimalSeparator = "."	' as used in amounts
Const MemoChunkLength = 0	' if memo field consists of fixed length chunks
Const TxnDatePattern = ".*(\d\d)\.(\d\d)\.(\d\d)\ (\d\d)\.(\d\d)"	' pattern to find transaction date in the memo
Dim TxnDateSequence: TxnDateSequence = Array(3,2,1,4,5,0)	' order of the info in the pattern: Y,M,D,H,M,S
Const PayeeLocation = 0		' start of payee in memo
Const PayeeLength = 0		' length of payee in memo
Dim ThisYear: ThisYear = Year(Now())
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
' Date, Transactions, Credits, Debits, Balance
' more recently Credits and Debits have swapped places (this is handled in the code)
' Date, Transactions, Debits, Credits, Balance
Dim aFields
aFields = Array( _
	Array(fldBookDate, "Date"), _
	Array(fldMemo, "Transactions"), _
	Array(fldAmtCredit, "Credits"), _
	Array(fldAmtDebit, "Debits"), _
	Array(fldClosingBal, "Balance"), _
	Array(fldSkip, "") _
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
aPropertyList = Array()

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

' Sub Configure
'	If ShowConfigDialog(ScriptName, aPropertyList) Then
'		SaveProperties ScriptName, aPropertyList
'	End If
' End Sub

Function ParseDate(sDate)
' 20080114 CS: Dates seem now to look like "December 27" with no year!
	If IsNumeric(Left(sDate, 1)) Then
		ParseDate = ParseDateEx(sDate, DateSequence, DateSeparator)
	Else
		ParseDate = ParseDateEx(sDate & " " & CStr(ThisYear), "MDY", " ")
		If ParseDate <> NODATE Then
			If ParseDate > Now() Then
				ParseDate = DateAdd("yyyy", -1, ParseDate)
			End If
		End If
	End If
End Function

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
	For i=1 To SkipHeaderLines
		If AtEOF() Then
			Exit Function
		End If
		sLine = ReadLine()
		If StartsWith(sLine, "Date") Then
			Exit For
		End If
	Next
	If AtEOF() Then
		Exit Function
	End If
	If Not StartsWith(sLine, "Date") Then
		sLine = ReadLine()
	End If
	If CSVSeparator = "" Then
		vFields = ParseLineFixed(sLine, TxnLinePattern)
	Else
		vFields = ParseLineDelimited(sLine, CSVSeparator)
	End If
	If TypeName(vFields) <> "Variant()" Then
		DebugMessage "not var array"
		Exit Function
	End If
	If UBound(vFields) < MinFieldsExpected Or UBound(vFields) > MaxFieldsExpected Then
		DebugMessage "Wrong number of fields - got " & UBound(vFields) & ", expected " _
			& MinFieldsExpected & "-" & MaxFieldsExpected & " - " & sLine
		Exit function
	End If
' check for reversed debits and credits
	If vFields(3) = "Debits" Then
		DebugMessage "Swapping debits and credits"
		i = aFields(2)
		aFields(2) = aFields(3)
		aFields(3) = i
	End If
	If ColumnHeadersPresent Then
		For i=1 To UBound(vFields)
			If UBound(aFields(i-1)) > 1 Then
				If Not StringMatches(vFields(i), aFields(i-1)(2)) Then
					DebugMessage "Field " & CStr(i) & ": '" & vFields(i) & "' does not match '" & aFields(i-1)(2) & "'"
					Exit Function
				End If
			Else
				If Trim(vFields(i)) <> aFields(i-1)(1) Then
					DebugMessage "Field " & CStr(i) & " " & aFields(i-1)(1) & " instead of " & vFields(i)
					Exit function
				End If
			End If
		Next
	Else
' pattern-match the first row
		For i=1 To UBound(vFields)
			sField = Trim(vFields(i))
			If UBound(aFields(i-1)) > 1 Then
				sPat = aFields(i-1)(2)
				bTmp = StringMatches(sField, sPat)
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
				DebugMessage "Field " & i & " (" & sField & ") failed to match '" & sPat & "'"
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
	Dim iTmp		' temporary counter
	Dim vDateBits	' parts of date
	Dim iSeq		' transaction sequence number
	Dim i
	Dim dBal		' temp balance date
	Dim sField		' field value being processed
	Dim dAvailBal, dClosingBal, sBalDate
	Dim dLastBookDate, iTxnSeq	' for generating FITIDs
   Dim da: Set da = New DateAccumulator

	LoadTextFile = False
	sAcct = ""
'Account name: ,accountname,,,,
'Account balance: ,£1084.21,,,,
'Available balance: ,£1036.05,,,,
    For i=1 To SkipHeaderLines
        sLine = ReadLine()
        vFields = ParseLineDelimited(sLine, CSVSeparator)
        
        If TypeName(vFields) = "Variant()" Then
            If UBound(vFields) >= 2 Then
                Select Case LCase(vFields(1))
                Case "available balance: "
                    dAvailBal = ParseNumber(vFields(2), DecimalSeparator)
                Case "account balance: "
                    dClosingBal = ParseNumber(vFields(2), DecimalSeparator)
                Case "account : "
                    sAcct = Trim(vFields(2))
                Case "account name: "
                    sAcct = Trim(vFields(2))
                Case "date"
                    Exit For
                End Select
            Else
                If StartsWith(sLine, "Available Balance: ") or StartsWith(sLine, "Available balance: ") Then
                    dAvailBal = ParseNumber(Mid(sLine, 20), DecimalSeparator)
                ElseIf StartsWith(sLine, "Account Balance: ") Or StartsWith(sLine, "Account balance: ") Then
                    dClosingBal = ParseNumber(Mid(sLine, 18), DecimalSeparator)
                ElseIf StartsWith(sLine, "Account : ") Then
                    sAcct = Trim(Mid(sLine, 11))
                ElseIf StartsWith(sLine, "Account name: ") Then
                    sAcct = Trim(Mid(sLine, 16))
                    iTmp = InStr(sAcct, ",")
                    If iTmp > 0 Then sAcct = Trim(Left(sAcct, iTmp-1))
                ElseIf StartsWith(sLine, "Date") Then
                    Exit For
                End If
            End If
        Else
            If StartsWith(sLine, "Available Balance: ") or StartsWith(sLine, "Available balance: ") Then
                dAvailBal = ParseNumber(Mid(sLine, 20), DecimalSeparator)
            ElseIf StartsWith(sLine, "Account Balance: ") Or StartsWith(sLine, "Account balance: ") Then
                dClosingBal = ParseNumber(Mid(sLine, 18), DecimalSeparator)
            ElseIf StartsWith(sLine, "Account : ") Then
                sAcct = Trim(Mid(sLine, 11))
            ElseIf StartsWith(sLine, "Account name: ") Then
                sAcct = Trim(Mid(sLine, 16))
                iTmp = InStr(sAcct, ",")
                If iTmp > 0 Then sAcct = Trim(Left(sAcct, iTmp-1))
            ElseIf StartsWith(sLine, "Date") Then
                Exit For
            End If
        End If
    Next
	If ColumnHeadersPresent Then
		If Not StartsWith(sLine, "Date") Then
			sLine = ReadLine()
		End If
	End if
	Do While Not AtEOF()
		sLine = ReadLine()
		If Len(sLine) > 0 And Left(sLine,1) <> CSVSeparator Then
			If CSVSeparator = "" Then
				vFields = ParseLineFixed(sLine, TxnLinePattern)
			Else
				vFields = ParseLineDelimited(sLine, CSVSeparator)
			End If
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage, vbOkOnly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If
			If UBound(vFields) < MinFieldsExpected Or UBound(vFields) > MaxFieldsExpected Then
				DebugMessage True, True, "Wrong number of fields - " & CStr(UBound(vFields)) & " - " & sLine
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
					Stmt.BranchName = BranchCode
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
					Stmt.Acct = sAcct
					Stmt.BranchName = BranchCode
				End If
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
					Stmt.AvailableBalance.Ccy = sField
				case fldClosingBal
					Stmt.ClosingBalance.Amt = ParseNumber(sField, DecimalSeparator)
				case fldAvailBal
					Stmt.AvailableBalance.Amt = ParseNumber(sField, DecimalSeparator)
				case fldBookDate
					Txn.BookDate = ParseDate(sField)
               da.Process(Txn.BookDate)
				case fldValueDate
					Txn.ValueDate = ParseDate(sField)
				Case fldTransactionDate
					Txn.TxnDate = ParseDate(sField)
					Txn.TxnDateValid = (Txn.TxnDate <> NODATE)
               da.Process(Txn.TxnDate)
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
			
			Dim sMemo, sMemoLower
' find the payee, transaction type and txn date if we can

' Nationwide: locate the payee in the memo field
			sMemo = Txn.Memo
        If Right(sMemo, 1) = "." Then
            sTmp = Left(sMemo, Len(sMemo)-1)
        Else
            iTmp = InStr(sMemo, ".")
            If iTmp>0 Then
                sTmp = Trim(Left(sMemo, iTmp-1))
            Else
                sTmp = Trim(sMemo)
            End If
        End If
			Txn.Payee = sTmp
         sMemoLower = LCase(sMemo)
			If StartsWith(sMemoLower, "cheque credit") Then
				Txn.TxnType = "CREDIT"
			ElseIf StartsWith(sMemoLower, "cash credit") Then
				Txn.TxnType = "CREDIT"
         ElseIf StartsWith(sMemoLower, "cheque credit") Then
            Txn.TxnType = "CHECK"
			ElseIf StartsWith(sMemoLower, "cheque") Then
				Txn.CheckNum = Mid(sTmp, 8)
				Txn.TxnType = "CHECK"
			ElseIf StartsWith(sMemoLower, "transfer to") Then
				Txn.Payee = Trim(Mid(sTmp, 13))
			ElseIf StartsWith(sMemoLower, "transfer from") Then
				Txn.Payee = Trim(Mid(sTmp, 14))
			ElseIf StartsWith(sMemoLower, "standing order") Then
				Txn.Payee = Trim(Mid(sTmp, 15))
			ElseIf StartsWith(sMemoLower, "cash machine wdl") _
				Or StartsWith(sMemoLower, "branch or cash machine withdrawal") Then
				Txn.Payee = "Cash Withdrawal"
				Txn.TxnType = "ATM"
			ElseIf StartsWith(sMemoLower, "direct debit") Then
				Txn.Payee = Trim(Mid(sTmp, 14))
				Txn.TxnType = "DIRECTDEBIT"
			ElseIf StartsWith(sMemoLower, "bank credit") Then
				Txn.Payee = Trim(Mid(sTmp, 13))
				Txn.TxnType = "DIRECTDEP"
			ElseIf StartsWith(sMemoLower, "interest") Then
				Txn.TxnType = "INT"
			ElseIf StartsWith(sMemoLower, "correction") Then
				Txn.TxnType = "OTHER"
			End If

' Nationwide: extract transaction dates			
			vDateBits = ParseLineFixed(sMemo, ".*Withdrawal [Dd]ate (\d+ .+ \d+).*")
         If VersionAtLeast("3.5.36") Then
            If UBound(vDateBits) < 0 Then
                vDateBits = ParseLineFixed(sMemo, ".*Credited [Oo]n (\d+ .+ \d+).*")
            End If
            If UBound(vDateBits) > 0 Then
                Txn.TxnDate = ParseDate(vDateBits(1))
                Txn.TxnDateValid = True
            End If
         Else
            If TypeName(vDateBits) <> "Variant()" Then
                vDateBits = ParseLineFixed(sMemo, ".*Credited [Oo]n (\d+ .+ \d+).*")
            End If
            If TypeName(vDateBits) = "Variant()" Then
                Txn.TxnDate = ParseDate(vDateBits(1))
                Txn.TxnDateValid = True
            End If
         End If

' Nationwide: sort out an ID based on the book date (default is statement date, which won't work in this case because
' download periods can overlap)
			If Txn.BookDate = dLastBookDate Then
				iTxnSeq = iTxnSeq + 1
			Else
				iTxnSeq = 1
			End If
			Txn.FITID = CStr(Year(Txn.BookDate)) & "." & Right("00" & CStr(DatePart("y", Txn.BookDate)), 3) & "." & CStr(iTxnSeq)
			dLastBookDate = Txn.BookDate

' tidy up the memo
			If MemoChunkLength > 0 Then
				sMemo = Txn.Memo
				Txn.Memo = ""
				For i=1 To Len(sMemo) Step MemoChunkLength
					ConcatMemo Trim(Mid(sMemo, i, MemoChunkLength))
				Next
			End If

' keep tabs on the statement/balance Date
			Stmt.ClosingBalance.BalDate = da.MaxDate
			Stmt.OpeningBalance.BalDate = da.MinDate
		End If
	Loop
	sTmp = MapAccount(Stmt.Acct)
	If Len(sTmp) > 0 Then Stmt.Acct = sTmp
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
		dTxn = DateSerial(iYear, iMonth, iDay) + TimeSerial(iHour, iMin, iSec)
	End If
	TransDate = dTxn
End Function

Sub DebugMessage(sLine)
	If DebugRecognition Then
		Message True, True, sLine, ScriptName
	End If
End Sub
