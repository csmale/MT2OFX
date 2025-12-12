' MT2OFX Input Processing Script Egg Bank HTML format

Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/EggUK-HTML.vbs 2     6/10/09 0:07 Colin $"

Dim sFile	' to hold the whole file
Dim iPtr	' pointer in this big string

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
' Egg uses xhtml so we can treat it as XML
Function RecogniseTextFile()
	RecogniseTextFile = False
	On Error Resume Next
	Dim sFile: sFile = EntireFileHTML()
	If InStr(sFile, "egg.com") = 0 Then
'		MsgBox "Not from egg.com"
		Exit Function
	End If
	Dim iTmp
' find transaction table
	iTmp = InStr(sFile, "<table id=""tblTransactionsTable""")
	If iTmp = 0 Then
'		MsgBox "Cannot find start of table"
		Exit Function
	End If
'	MsgBox "found transactions: " & Mid(sFile, iTmp, 100)
	RecogniseTextFile = True
	
	If RecogniseTextFile Then
		LogProgress ScriptName, "File Recognised"
	End If
End Function

Function ExtractValueByClass(sTxn, sClass)
	Dim re, m
	Set re = New RegExp
	re.Pattern = Replace("<td class=""%1"">(.*?)</td>", "%1", sClass)
	re.Global = False
	Set m=re.Execute(sTxn)
	If m.Count > 0 Then
		ExtractValueByClass = m(0).SubMatches(0)
	Else
		ExtractValueByClass = ""
	End If
End Function

Function ParseAmount(sAmt)
	Dim bDR
	bDR = InStr(sAmt, "DR")
	
	ParseAmount = ParseNumber(Replace(sAmt, "DR", ""), ".")
	If bDR Then
		ParseAmount = -ParseAmount
	End If
'	MsgBox "amount: " & samt & " = " & ParseAmount
End Function

Function LoadTextFile()
	LoadTextFile = False
	Dim iTmp, iTmp2
	sFile = EntireFileHTML()
	Dim sPart
	Dim re, m, m1, sTxn
	Dim reTxn, mTxn
	Dim sAmt, sDescr, sDate, sCat
	Dim sBal, sTmp, sTmp2
	Dim Stmt, sAcct

' we can reconstruct the whole credit card number from this lot:
'	<div class="input">
'		<label for="lblCardNumber">account number</label>
'		<span id="lblCardNumber" class="creditcard">5186 1111 2222 XXXX</span>
'		<span id="lblCardNumberForPrint" class="creditcardforprint">5186 **** **** 3333</span>
'	</div>
	
	Set re = New RegExp
	re.Pattern = "creditcard"">(.+?)</span>"
	re.Global = False
	re.IgnoreCase = False
	Set m=re.Execute(sFile)
	If m.Count > 0 Then
		sTmp = Replace(m(0).SubMatches(0), " ", "")
	End If
	re.Pattern = "creditcardforprint"">(.+?)</span>"
	Set m=re.Execute(sFile)
	If m.Count > 0 Then
		sTmp2 = Replace(m(0).SubMatches(0), " ", "")
	End If
	sAcct = Left(sTmp, 12) & Right(sTmp2, 4)
	
' find transaction table
	iTmp = InStr(sFile, "<table id=""tblTransactionsTable""")
	If iTmp = 0 Then
		MsgBox "Cannot find start of table"
		Exit Function
	End If
	iPtr = iTmp
	
'	MsgBox "found transactions: " & Mid(sFile, iTmp, 100)
' extract just the table we need
	iTmp2 = InStr(iTmp, sFile, "<div class=""calltoaction""")
	If iTmp2 > 0 Then
		sPart = Mid(sFile, iTmp, iTmp2-iTmp)
	Else
		sPart = Mid(sFile, iTmp)
	End If

	Set Stmt = NewStatement()
	Stmt.Acct = sAcct
	Stmt.AcctType = "CREDITCARD"
	Stmt.OpeningBalance.BalDate = NODATE
	Stmt.OpeningBalance.Ccy = "GBP"
	Stmt.AvailableBalance.BalDate = NODATE
	Stmt.ClosingBalance.BalDate = NODATE				
	Stmt.ClosingBalance.Ccy = "GBP"
	Stmt.BankName = "EGGGGB2L"
	Stmt.BranchName = ""

' get the balance from the header of the transaction table
	re.Pattern = ".*closing balance:</th>\s+<td[^>]*>(.+?)</td.*"
	re.Global = False
	re.IgnoreCase = False
	set m = re.Execute(sPart)
	If m.Count <> 0 Then
		sBal = m(0).SubMatches(0)
' MsgBox "balance found: " & sBal
		Stmt.ClosingBalance.Amt = ParseAmount(sBal)
	End If

' get the transactions so we can handle one at a Time
	iTmp = InStr(sPart, "<tbody>")
	If iTmp > 0 Then
		sPart = Mid(sPart, iTmp+8)
	Else
		Exit Function
	End If
	
	re.Pattern = "<tr[^>]*>((.|\s)+?)</tr>"
	re.Global = True
	Set m = re.Execute(sPart)
	For Each m1 In m
		sTxn = m1.Value
		sDate = ExtractValueByClass(sTxn, "date")
		sDescr = ExtractValueByClass(sTxn, "description")		
		sCat = ExtractValueByClass(sTxn, "category")
		sAmt = ExtractValueByClass(sTxn, "money")
'		MsgBox "Transaction: " & sDescr		
		If sDate <> "" Then
			NewTransaction()
			Txn.BookDate = ParseDateEx(sDate, "DMY", " ")
			If Txn.BookDate = NODATE Then
				MsgBox "Bad date '" & sDate & "': " & ParseDateError
			End if
			Txn.Memo = sDescr
			Txn.Payee = Trim(Left(sDescr, 24))
			Txn.Amt = ParseAmount(sAmt)
			If Txn.Amt < 0 Then
				Txn.TxnType = "PAYMENT"
			Else
				Txn.TxnType = "DEP"
			End If
			If Stmt.OpeningBalance.BalDate = NODATE Then
				Stmt.OpeningBalance.BalDate = Txn.BookDate
			End If
			Stmt.ClosingBalance.BalDate = Txn.BookDate
		Else
			If Not (Txn Is Nothing) Then
				ConcatMemo sDescr
			End If
		End If
	Next
	
'	    <caption>Transaction details</caption>
'	    <thead>
'		    <tr>
'			    <th class="date">date</th>
'			    <th class="description">description</th>
'			    <th class="category">category</th>
'			    <th class="amount">amount</th>
'		    </tr>
'	    </thead>

'        <tfoot>
'		    <tr>
'			    <th colspan="3">closing balance:</th>
'			    <td class="money">£1,331.69 DR</td>
'		    </tr>
'		    <tr>
'			    <th colspan="3">minimum payment:</th>
'			    <td class="money">£26.63</td>
'		    </tr>
'	    </tfoot>	    

' after end of table:
'    <div class="calltoaction sectionend">

	LoadTextFile = True
End Function
