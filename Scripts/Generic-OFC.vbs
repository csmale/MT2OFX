' MT2OFX Input Processing Script for OFC files

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/Generic-OFC.vbs 19    12/07/09 12:13 Colin $"

Const ScriptName = "Generic-OFC"
Const FormatName = "OFC File Format"
Const ParseErrorMessage = "Parse Error"
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

' Property List is an array of arrays, each of which has the following elements:
'	1. Property key - used to reference properties
'	2. Property name - used as a label in the config screen
'	3. Property description - used as a description or tooltip in the config screen
'	4. Data type - ptString, ptBoolean, ptInteger, ptFloat, ptDate, ptChoose
'	5. Value list (will be displayed in a combobox) - array of values (Only with ptChoose)
' 20080129 CS: Added OFCIgnoreZeroTxns and OFCInvertBalance
Dim aPropertyList
aPropertyList = Array( _
	Array("OFCCurrency", "OFC Currency", _
		"The ISO currency code assumed for OFC input. Must be 3 letters, e.g. USD, EUR, GBP.", _
		ptCurrency), _
	Array("OFCNewFITID", "Force new FITID", _
		"Set to True if the OFC input file does not contain reliable unique transaction identifiers.", _
		ptBoolean), _
	Array("OFCQuickenBankID", "Quicken Bank ID", _
		"This value will be used as the value of INTU.BID in a QFX file. A value of 0 will not appear in the output.", _
		ptInteger), _
	Array("OFCZeroTxnTreatment", "Treatment of zero transactions", _
		"Transactions with an amount of 0 can be processed normally, combined (MEMO/NAME) with the preceding transaction or will not appear in the output.", _
		ptChoice, Array("Include", "Combine", "Ignore")), _
	Array("OFCInvertBalance", "Invert balance", _
		"If checked, the sign of the account balance is inverted.", _
		ptBoolean), _
	Array("OFCSwapMemoAndName", "Swap memo and name", _
		"If checked, the MEMO and NAME fields are swapped over.", _
		ptBoolean) _
	)

Dim sXML		' intermediate XML format string
Dim xDoc		' XML document being processed
Dim ElementsDict	' Dictionary for OFC elements

' OFC does not include an explicit currency code! Change the script
' property through the Options screen to reflect the currency of your account
Dim CurrencyCode

Sub Initialise()
    LogProgress ScriptName, "Initialise"
    Set xDoc = Nothing
	If Not CheckVersion() Then
		Abort
	End If
	LoadProperties ScriptName, aPropertyList
End Sub

Sub Configure
	If ShowConfigDialog(ScriptName, aPropertyList) Then
		SaveProperties ScriptName, aPropertyList
	End If
End Sub

' function DescriptiveName
' returns a string with a descriptive name of this script
Function DescriptiveName()
	DescriptiveName = "For processing OFC files as input for the benefit of users of Money 2005."
End Function

Function StartsWith(s, Prefix)
	StartsWith = (Left(s,Len(Prefix)) = Prefix)
End Function

' handle dates/times in the formats YYYYMMDDHHMMSS and YYYYMMDD
Function ParseDate(sDate)
	Dim iYear, iMonth, iDay			' for dates
	Dim iHour, iMin, iSec			' for times
	iYear = CInt(Left(sDate,4))
	iMonth = CInt(Mid(sDate,5,2))
	iDay = CInt(Mid(sDate,7,2))
	ParseDate = DateSerial(iYear, iMonth, iDay)
	If Len(sDate) > 8 Then
		iHour = CInt(Mid(sDate,9,2))
		iMin = CInt(Mid(sDate,11,2))
		iSec = CInt(Mid(sDate,13,2))
		ParseDate = ParseDate + TimeSerial(iHour, iMin, iSec)
	End If
End Function

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
	Dim sLine
	Dim i
	Dim bRet: bRet = False
	For i=1 To 10
		If Not AtEOF() Then
			sLine = Replace(ReadLine(), vbTab, "")
			If Trim(sLine) = "<OFC>" Then
				bRet = True
				Exit For
			End If
		End If
	Next
	If bRet Then
		LogProgress ScriptName, "File Recognised"
	End If
	If bRet Then
		MakeElementList
		bRet = LoadAsXML()	' leaves xDoc populated
	End If
	RecogniseTextFile = bRet
End Function

' Construct list of tags which start an OFC Element. These tags have no
' corresponding closing tags in OFC which must be added! Other tags in the
' OFC file should be closed properly.
Sub MakeElementList()
	Dim aElements	' list of tage which represent OFC Elements
	aElements = Array( _
		"ACCTID", _
		"ACCTTYPE", _
		"ACTION", _
		"ADDRESS", _
		"BANKID", _
		"BRANCHID", _
		"CHKNUM", _
		"CITY", _
		"CLTID", _
		"CPAGE", _
		"DAYSREQD", _
		"DAYSWITH", _
		"DTCLIENT", _
		"DTD", _
		"DTDUE", _
		"DTEND", _
		"DTPOSTED", _
		"DTSERVER", _
		"DTSTART", _
		"ERROR", _
		"FITID", _
		"LEDGER", _
		"MEMO", _
		"NAME", _
		"NEWPASS", _
		"PAYACCT", _
		"PAYEEID", _
		"PHONE", _
		"POSTALID", _
		"SERVICE", _
		"SESSKEY", _
		"SIC", _
		"SRVRTID", _
		"STATE", _
		"STATUS", _
		"TRNAMT", _
		"TRNTYPE", _
		"USERID", _
		"USERPASS" _
	)

	Set ElementsDict = CreateObject("Scripting.Dictionary")
	Dim e
	For Each e In aElements
		ElementsDict.Add e,""
	Next
End Sub

' checks to see if the given tag is an OFC Element (see above)
Function IsOFCElement(sTag)
	IsOFCElement = ElementsDict.Exists(sTag)
End Function

' sort out XML entity encoding which may or may not be present in the input text
Function FixXMLEncoding(sTextIn)
	Dim sText
	sText = Replace(sTextIn, "&amp;", "&")
	sText = Replace(sText, "&lt;", "<")
	sText = Replace(sText, "&", "&amp;")
	sText = Replace(sText, "<", "&lt;")
	FixXMLEncoding = sText
End Function

' translates the entire input file to an XML-compliant String
' OFC is almost there, but "Elements" in OFC-speak do not have a closing tag,
' which we have to add. "Aggregates" and "Records" do have a closing tag, so we
' don't add another one!
' also each line seems to start (apart from white space) with a tag!
' I hope this heuristic holds good, otherwise the routine below will need rewriting.
Function OFCtoXML()
	Dim sLine
	Dim iGT
	Dim iEndTag
	Dim sTag
	Dim sText, sEnd
	sXML = ""
	Rewind
	sLine = EntireFile()
	If Left(sLine, 5) = "<?xml" Then
' seems to be already XML compliant! except there MUST be a newline after the XML declaration!
		iGT = InStr(sLine, ">")
		sLine = Left(sLine, iGT) & VbCrLf & Mid(sLine, iGT+1)
		OFCtoXML = sLine
	Else
		Rewind
		Do While Not AtEOF()
			sEnd = ""
			sLine = Trim(Replace(ReadLine(),vbTab, ""))
			If Left(sLine, 1) = "<" Then
				If Left(sLine, 2) <> "</" Then
					iGT = InStr(sLine, ">")
					If iGT > 0 Then
						sTag = Mid(sLine, 2, iGT-2)
						If IsOFCElement(sTag) Then
							iEndTag = InStr(sLine, "</" & sTag & ">")
							if iEndTag = 0 Then
		' check for another tag at the end of the line. for now we don't actually process it,
		' just hope it produces valid XML...
								iEndTag = InStr(iGT, sLine, "<")
								If iEndTag = 0 Then
									sText = FixXMLEncoding(Mid(sLine, iGT+1))
								Else
									sText = FixXMLEncoding(Mid(sLine, iGT+1, iEndTag-iGT-1))
									sEnd = Mid(sLine, iEndTag)
' department of quick hacks: <PAYEE>...<MEMO>...
                            If Left(sEnd, 6) = "<MEMO>" Then sEnd = vbCrLf & "<MEMO>" & FixXMLEncoding(Mid(sEnd, 7)) & "</MEMO>"
								End If
							Else
								sText = FixXMLEncoding(Mid(sLine, iGT+1, iEndTag-iGT-1))
							End if
							sLine = "<" & sTag & ">" & sText & "</" & sTag & ">" & sEnd
						End If
					End If
				End If
			End If
			sXML = sXML & sLine & VbCrLf
		Loop
'		MsgBox "At EOF: " & sLine
'		Dim fso
'		Set fso=CreateObject("Scripting.FileSystemObject")
'		Dim f
'		Set f=fso.CreateTextFile("c:\xml.txt", True, True)
'		f.Write sXML
'		f.Close
		OFCtoXML = sXML
	End If
End Function

Function LoadAsXML()
	Dim bRet: bRet = False
	Dim sXML
	sXML = OFCtoXML()
	If xDoc Is Nothing Then
		Set xDoc = CreateObject("MSXML2.DOMDocument")
		xDoc.async = False
		xDoc.validateOnParse = False
	End If
	bRet = xDoc.loadXML(sXML)
	If bRet Then
		bRet=(xDoc.documentElement.NodeName = "OFC")
		If Not bRet Then
			MsgBox "Document is not an OFC file!",vbokonly+vbcritical,"OFC Processing Error"
		End If
	Else
		Dim myErr
		Set myErr = xDoc.parseError
		MsgBox myErr.reason _
			& "At line " & myErr.line & ", character " & myErr.linepos & vbcrlf _
			& "Text """ & myErr.srcText & """",vbokonly+vbcritical, "XML Load Error"
	End If
	LoadAsXML = bRet
End Function

Function NodeText(xNode)
	If xNode Is Nothing Then
		NodeText = ""
	Else
		NodeText = xNode.text
	End If
End Function

Function ParseAmount(sAmt)
	Dim iComma, iPoint
	iComma = InStr(sAmt, ",")
	iPoint = InStr(sAmt, ".")
	If iComma = 0 Then
		ParseAmount = ParseNumber(sAmt, ".")
	ElseIf iPoint = 0 Then
		ParseAmount = ParseNumber(sAmt, ",")
	Else	' both comma and point - assume the first is thousands and the last is decimal
		If iComma > iPoint Then	' point first
			ParseAmount = ParseNumber(sAmt, ",")		
		Else	' comma first
			ParseAmount = ParseNumber(sAmt, ".")		
		End If
	End If
End Function

' parse the input file, which is present in xDoc as an XMLDocument
Function LoadTextFile()
	Dim xStmts	' list of statements in the input
	Dim xStmt	' this statement
	Dim xTxns	' list of transactions in a statement
	Dim xTxn	' this transaction
	Dim Stmt	' output transaction
	Dim xAcct	' account aggregate
	Dim bNewFITID	' force new FITID
' 20080129 CS: Added OFCIgnoreZeroTxns and OFCInvertBalance
	Dim dAmt
	Dim sZeroTreatment	' Include, Ignore or Combine txns with amt=0
	Dim bIsFirstTxn		' shows first transaction in statement
	Dim bInvertBalance	' invert the sign of the balance
	Dim bSwapMemoAndName
	Dim sPayee, sMemo
	Dim i
	Dim j
	CurrencyCode = Ucase(GetProperty("OFCCurrency"))
	If Len(CurrencyCode) <> 3 Then
' 20080923 CS: Default to user's currency configured in Windows
		CurrencyCode = UserCurrency()
	End If
	If CurrencyCode = "" Then
		MsgBox "Please set the OFC Currency in the script properties through the Options screen before converting an OFC file.",vbOkOnly,"Unknown currency"
		Abort
		LoadTextFile = False
		Exit Function
	End If
	bNewFITID = CBool(GetProperty("OFCNewFITID"))
	If CStr(GetProperty("OFCQuickenBankID")) <> "0" Then
		Bcfg.IntuitBankID = CStr(GetProperty("OFCQuickenBankID"))
	End If
	sZeroTreatment = CStr(GetProperty("OFCZeroTxnTreatment"))
	bInvertBalance = CBool(GetProperty("OFCInvertBalance"))
	bSwapMemoAndName = CBool(GetProperty("OFCSwapMemoAndName"))

' If the recognition process was bypassed by a direct association we must parse
' the file fully now
	If xDoc Is Nothing Then
		MakeElementList
		If Not LoadAsXML() Then
			Abort
			LoadTextFile = False
			Exit Function
		End If
	End If

	Set xStmts = xDoc.selectNodes("/OFC/ACCTSTMT")
'	MsgBox "Statements found: " & xStmts.length
	For i=0 To xStmts.length-1
		Set xStmt = xStmts.item(i)
' process statement here!
		Set Stmt = NewStatement()
		bIsFirstTxn = True
		Dim sTmp
' at least one case has been seen of the account data enclosed in <ACCOUNT>..</ACCOUNT>!
		Set xAcct = xStmt.selectSingleNode("ACCTFROM")
		If Len(NodeText(xAcct.selectSingleNode("ACCOUNT"))) > 0 Then
			Set xAcct = xAcct.selectSingleNode("ACCOUNT")
		End If
' 20080102 CS: Process the ACCTTYPE!
		Select Case NodeText(xAcct.selectSingleNode("ACCTTYPE"))
			Case "0" : Stmt.AcctType = "CHECKING"
			Case "1" : Stmt.AcctType = "SAVINGS"
			Case "2" : Stmt.AcctType = "CREDITCARD"
			Case "3" : Stmt.AcctType = "MONEYMARKET"	' not supported
			Case "4" : Stmt.AcctType = "CREDITLINE"	    ' not supported
			Case Else : Stmt.AcctType = "OTHER"			' not supported
		End Select
		Stmt.BankName = NodeText(xAcct.selectSingleNode("BANKID"))
		Stmt.Acct = NodeText(xAcct.selectSingleNode("ACCTID"))
		Stmt.ClosingBalance.Amt = ParseAmount(NodeText(xStmt.selectSingleNode("STMTRS/LEDGER")))
		If bInvertBalance Then
			Stmt.ClosingBalance.Amt = -Stmt.ClosingBalance.Amt
		End If
		Stmt.ClosingBalance.Ccy = CurrencyCode
		Stmt.ClosingBalance.BalDate = ParseDate(NodeText(xStmt.selectSingleNode("STMTRS/DTEND")))
		Stmt.OpeningBalance.Ccy = ""
		Stmt.OpeningBalance.BalDate = ParseDate(NodeText(xStmt.selectSingleNode("STMTRS/DTSTART")))
		Set xTxns = xStmt.selectNodes("STMTRS/STMTTRN")
'		MsgBox "Transactions found: " & xTxns.length
		For j=0 To xTxns.length-1
			Set xTxn = xTxns.item(j)
' process transaction here!
' at least one case has been seen of the txn data enclosed in <GENTRN>..</GENTRN>!
			If Len(NodeText(xTxn.selectSingleNode("GENTRN"))) > 0 Then
				Set xTxn = xTxn.selectSingleNode("GENTRN")
			End If
			dAmt = ParseAmount(NodeText(xTxn.selectSingleNode("TRNAMT")))
			sPayee = NodeText(xTxn.selectSingleNode("NAME"))
			If Len(sPayee) = 0 Then
				sPayee = NodeText(xTxn.selectSingleNode("PAYEE/NAME"))
			End If
			sMemo = NodeText(xTxn.selectSingleNode("MEMO"))
			If dAmt <> 0 Or sZeroTreatment = "Include" Or bIsFirstTxn Then
				NewTransaction
				Txn.Amt = dAmt
				Txn.BookDate = ParseDate(NodeText(xTxn.selectSingleNode("DTPOSTED")))
'				Txn.ValueDate = Txn.BookDate	' OFC does not have Value Date
				Txn.Payee = sPayee
				Txn.Memo = sMemo
'				MsgBox Txn.Memo
' 20061029 CS: handle non-numeric TRNTYPE more gracefully
				sTmp = NodeText(xTxn.selectSingleNode("TRNTYPE"))
				If IsNumeric(sTmp) Then
					Txn.TxnType = OFXTxnType(CInt(sTmp))
				Else
					Txn.TxnType = "OTHER"
				End If
				Txn.FITID = NodeText(xTxn.selectSingleNode("FITID"))
				Txn.CheckNum = NodeText(xTxn.selectSingleNode("CHKNUM"))
				If bNewFITID Then
					Txn.FITID = ""	' let the main program invent one based on statement date
				End If
				If bSwapMemoAndName Then
					sTmp = Txn.Memo
					Txn.Memo = Txn.Payee
					Txn.Payee = sTmp
				End If
			ElseIf Not bIsFirstTxn and sZeroTreatment = "Combine" Then
				If bSwapMemoAndName Then
					sTmp = sMemo
					sMemo = sPayee
					sPayee = sTmp
				End If
				If Len(Txn.Payee) = 0 Then
					Txn.Payee = sPayee
				Else
					Txn.Payee = Txn.Payee & Cfg.MemoDelimiter & sPayee
				End If
				If Len(Txn.Memo) = 0 Then
					Txn.Memo = sMemo
				Else
					Txn.Memo = Txn.Memo & Cfg.MemoDelimiter & sMemo
				End If
			End If
			bIsFirstTxn = False
		Next
	Next
	LoadTextFile = True
End Function

Function OFXTxnType(iCode)
	Select Case iCode
	Case 0:		OFXTxnType = "CREDIT"
	Case 1:		OFXTxnType = "DEBIT"
	Case 2:		OFXTxnType = "INT"
	Case 3:		OFXTxnType = "DIV"
	Case 4:		OFXTxnType = "SRVCHG"
	Case 5:		OFXTxnType = "DEP"
	Case 6:		OFXTxnType = "ATM"
	Case 7:		OFXTxnType = "XFER"
	Case 8:		OFXTxnType = "CHECK"
	Case 9:		OFXTxnType = "PAYMENT"
	Case 10:	OFXTxnType = "CASH"
	Case 11:	OFXTxnType = "DIRECTDEBIT"
	Case Else:	OFXTxnType = "OTHER"
	End Select
End Function
