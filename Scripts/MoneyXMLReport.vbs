' MT2OFX Input Processing Script for Money XML Export Format

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/MoneyXMLReport.vbs 3     2/11/05 23:03 Colin $"

Const ScriptName = "MSMoney-XML"
Const FormatName = "Microsoft Money XML Report Format"
Const ParseErrorMessage = "Parse Error"
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

Dim xDoc		' XML document being processed
Dim aColumns()	' list of columns in detail records in file

'	<row rlttype=
'	4	column headers
'	6	transaction detail
'	8	group trailer
'	11	grand total
'	15	divider
'	16	group trailer
'	21	opening balance

'	<title rlttype=
'	0	title
'	1	account name
'	7	group header
'	26	?

'	<field rft=
'	0	text
'	1	number
'	3	Date
'	4	cleared status
'	8	amount

Sub Initialise()
    LogProgress ScriptName, "Initialise"
    Set xDoc = Nothing
	If Not CheckVersion() Then
		Abort
	End If
End Sub

' function DescriptiveName
' returns a string with a descriptive name of this script
Function DescriptiveName()
	DescriptiveName = FormatName
End Function

Function StartsWith(s, Prefix)
	StartsWith = (Left(s,Len(Prefix)) = Prefix)
End Function

Function ParseDate(sDate)
	Dim iYear, iMonth, iDay			' for dates
	iYear = CInt(Left(sDate,4))
	iMonth = CInt(Mid(sDate,5,2))
	iDay = CInt(Mid(sDate,7,2))
	ParseDate = DateSerial(iYear, iMonth, iDay)
End Function

Function TrimTrailingDigits(s)
	Dim r
	Set r=New regexp
	r.Global = False
	r.Pattern = "^(.*?) *\d+$"
	Dim m
	Set m=r.Execute(s)
	If m.Count = 0 Then
		TrimTrailingDigits = s
	Else
		TrimTrailingDigits = m(0).SubMatches(0)
	End If
End Function

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
	Dim bRet
	Dim xNode
	Dim xFields
	Dim i
	If xDoc Is Nothing Then
		Set xDoc = CreateObject("MSXML2.DOMDocument")
		xDoc.async = False
	End If
	bRet = xDoc.load(Session.FileIn)
	If bRet Then
		bRet=(xDoc.documentElement.NodeName = "MoneyReport")
		If Not bRet Then
			MsgBox "Document is not a MoneyReport"
		End If
	End If
	If bRet then
		Set xNode = xDoc.selectSingleNode("//row[@rlttype='4']")
		If xNode Is Nothing Then
			MsgBox "No Column Headers Found (1)"
			bRet = False
		Else
			Set xFields = xNode.selectNodes("field")
			If xFields Is Nothing Then
				MsgBox "No Column Headers Found (2)"
				bRet = False
			Else
				If xFields.length = 0 Then
					bRet = False
					MsgBox "No Column Headers Found (3)"
				Else			
					MsgBox CStr(xFields.length) & " Column Headers Found"
					ReDim aColumns(xFields.length-1)
					For i=0 To xFields.length-1
						aColumns(i) = Trim(xFields.item(i).text)
					Next
				End If
			End If
		End If
	End If
	If bRet Then
		LogProgress ScriptName, "File Recognised"
	End If
	RecogniseTextFile = bRet
End Function

Function LoadTextFile()
	Dim xTxns	' list of transactions to be processed
	Dim xTxn	' transaction being processed
	Set xTxns = xDoc.selectNodes("//row[@rlttype='6']")
	Dim xCols	' columns in transaction
	Dim xCol	' individual column of transaction
	Dim i, j
	Dim sCol, sFldNum, iFldNum, sVal
	Dim Stmt
		
	Set Stmt = NewStatement()
	Stmt.Acct = "unknown"
	Stmt.BankName = "unknown"
	Stmt.StatementID = MakeGUID()
	
	MsgBox CStr(xTxns.length-1) & " Transactions"
	For i=0 To xTxns.length-1
		Set xTxn = xTxns.item(i)
		Set xCols = xTxn.selectNodes("field")
		NewTransaction
		For j=0 To xCols.length-1
			Set xCol = xCols.item(j)
			sFldNum = xCol.attributes.getNamedItem("num").text
			iFldNum = CInt(sFldNum)
			If iFldNum<0 Or iFldNum>UBound(aColumns) Then
				MsgBox "Field number out of bounds"
				LoadTextFile = False
				Exit Function
			End If
			sCol = aColumns(iFldNum)
			sVal = Trim(xCol.text)
			Select Case sCol
			Case "Amount"
				Txn.Amt = ParseNumber(sVal, ".")
			Case "Account"
			'	this had better stay the same...
			Case "C"
			'	cleared status - no action
			Case "Category"
				Txn.Category = sVal
			Case "Num"	'	check number
				Txn.CheckNum = sVal
			Case "Date"
			'	Txn.TxnDate = ParseDate(sVal)
			Case "Payee"
				Txn.Payee = sVal
			Case "Memo"
				Txn.Memo = sVal
			Case "VAT", "VAT %", "Net"
			'	ignore vat stuff
			Case "Running Balance"
			'	ignore for now
			Case Else
				MsgBox "Unexpected Column: " & sCol
			End Select
		Next
	Next
	LoadTextFile = True
End Function

