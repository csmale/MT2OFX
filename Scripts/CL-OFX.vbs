' MT2OFX Input Processing Script for OFX files

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/CL-OFX.vbs 1     25/11/08 23:57 Colin $"

Const ScriptName = "CL-OFX"
Const FormatName = "Crédit Lyonnais OFX File Format"
Const ParseErrorMessage = "Parse Error"
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

' Property List is an array of arrays, each of which has the following elements:
'	1. Property key - used to reference properties
'	2. Property name - used as a label in the config screen
'	3. Property description - used as a description or tooltip in the config screen
'	4. Data type - ptString, ptBoolean, ptInteger, ptFloat, ptDate, ptChoose
'	5. Value list (will be displayed in a combobox) - array of values (Only with ptChoose)
' Dim aPropertyList
' aPropertyList = Array( _
' 	Array("OFCCurrency", "OFC Currency", _
' 		"The ISO currency code assumed for OFC input. Must be 3 letters, e.g. USD, EUR, GBP.", _
' 		ptString) _
' 	)

Dim sXML		' intermediate XML format string
Dim xDoc		' XML document being processed
Dim ElementsDict	' Dictionary for OFX elements

Sub Initialise()
    LogProgress ScriptName, "Initialise"
    Set xDoc = Nothing
	If Not CheckVersion() Then
		Abort
	End If
'	LoadProperties ScriptName, aPropertyList
End Sub

'Sub Configure
'	If ShowConfigDialog(ScriptName, aPropertyList) Then
'		SaveProperties ScriptName, aPropertyList
'	End If
'End Sub

' function DescriptiveName
' returns a string with a descriptive name of this script
Function DescriptiveName()
	DescriptiveName = "For processing OFX files from Crédit Lyonnais."
End Function

Function StartsWith(s, Prefix)
	StartsWith = (Left(s,Len(Prefix)) = Prefix)
End Function

' handle dates/times in the formats YYYYMMDDHHMMSS and YYYYMMDD
Function ParseDate(sDate)
	Dim iYear, iMonth, iDay			' for dates
	Dim iHour, iMin, iSec			' for times
	If Len(sDate) < 8 Then
		ParseDate = DateSerial(0,0,0)
		Exit Function
	End If
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
	Dim bRet: bRet = False
	If Not AtEOF() Then
		sLine = Trim(Replace(ReadLine(), vbTab, ""))
		If StartsWith(sLine, "OFXHEADER:") Then
			bRet = True
		' AF->Credit Lyonnais : file begin with Content-Type
		ElseIf sLine = "Content-Type: application/x-ofx" Then
			bRet = True
		ElseIf sLine = "<OFX>" Then
			bRet = True
		Elseif StartsWith(sLine, "<?xml") Then
			bRet = True	' might be OFX2. LoadAsXML will check this out
		End If
	End If
	If bRet Then
		MakeElementList
		bRet = LoadAsXML()	' leaves xDoc populated
	End If
	If bRet Then
		LogProgress ScriptName, "File Recognised"
	End If
	RecogniseTextFile = bRet
End Function

' Construct list of tags which start an OFX Element. These tags have no
' corresponding closing tags in OFX which must be added! Other tags in the
' OFX file should be closed properly.
Sub MakeElementList()
	Dim aElements	' list of tage which represent OFX Elements
	aElements = Array( _
		"ACCRDINT", _
		"ACCTBAL", _
		"ACCTEDITMASK", _
		"ACCTFORMAT", _
		"ACCTID", _
		"ACCTKEY", _
		"ACCTREQUIRED", _
		"ACCTTYPE", _
		"ACCTTYPE2", _
		"ACTIVITY", _
		"ADDR1", _
		"ADDR2", _
		"ADDR3", _
		"ADJAMT", _
		"ADJDATE", _
		"ADJDESC", _
		"ADJNO", _
		"AMTDUE", _
		"APPID", _
		"APPVER", _
		"ASSETCLASS", _
		"AUCTION", _
		"AVAILACCTS", _
		"AVAILCASH", _
		"AVGCOSTBASIS", _
		"BALAMT", _
		"BALCLOSE", _
		"BALDNLD", _
		"BALMIN", _
		"BALOPEN", _
		"BALTYPE", _
		"BANKBRANCH", _
		"BANKCITY", _
		"BANKID", _
		"BANKNAME", _
		"BANKPOSTALCODE", _
		"BILLDETAILTABLETYPE", _
		"BILLERID", _
		"BILLERINFOURL", _
		"BILLID", _
		"BILLPMTSTATUSCODE", _
		"BILLPUB", _
		"BILLREFINFO", _
		"BILLSTATUSCODE", _
		"BILLTYPE", _
		"BOOKINGTEXT", _
		"BRANCHID", _
		"BRAND", _
		"BROKERID", _
		"BUSNAMEACCTHELP", _
		"BUYPOWER", _
		"BUYTYPE", _
		"C", _
		"CALLPRICE", _
		"CALLTYPE", _
		"CANADDPAYEE", _
		"CANBILLPAY", _
		"CANCELWND", _
		"CANEMAIL", _
		"CANMODMDLS", _
		"CANMODPMTS", _
		"CANMODSTATUS", _
		"CANMODXFERS", _
		"CANMOTO", _
		"CANMULTI", _
		"CANNOTIFY", _
		"CANPENDING", _
		"CANRECUR", _
		"CANSCHED", _
		"CANSUPPORTGROUPID", _
		"CANSUPPORTIMAGES", _
		"CANSUPPORTUSERID", _
		"CANUPDATEPRESNAMEADDRESS", _
		"CANUSEDESC", _
		"CANUSERANGE", _
		"CASESEN", _
		"CHARTYPE", _
		"CHE.PTTACCTID", _
		"CHECKING", _
		"CHECKNUM", _
		"CHGPINFIRST", _
		"CHGUSERINFO", _
		"CHKANDDEB", _
		"CHKERROR", _
		"CHKNUMEND", _
		"CHKNUMSTART", _
		"CHKSTATUS", _
		"CITY", _
		"CLIENTACTREQ", _
		"CLOSINGAVAIL", _
		"CLTCOOKIT", _
		"CODE", _
		"COLNAME", _
		"COLTYPE", _
		"COMMISSION", _
		"CONFMSG", _
		"CONSUPOSTALCODE", _
		"CORRECTACTION", _
		"CORRECTFITID", _
		"COUNT", _
		"COUNTRY", _
		"COUPONFREQ", _
		"COUPONRT", _
		"CREDITLIMIT", _
		"CSPHONE", _
		"CURDEF", _
		"DATEBIRTH", _
		"DAYPHONE", _
		"DAYSTOPAY", _
		"DAYSWITH", _
		"DEBADJ", _
		"DEBTCLASS", _
		"DEBTTYPE", _
		"DENOMINATOR", _
		"DEPANDCREDIT", _
		"DESC", _
		"DETAILAVAILABLE", _
		"DFLTDAYSTOPAY", _
		"DIFFFIRSTPMT", _
		"DIFFLASTPMT", _
		"DOMXFERFEE", _
		"DSCAMT", _
		"DSCDATE", _
		"DSCDESC", _
		"DSCRATE", _
		"DTACCTUP", _
		"DTASOF", _
		"DTAUCTION", _
		"DTAVAIL", _
		"DTBILL", _
		"DTCALL", _
		"DTCHANGED", _
		"DTCLIENT", _
		"DTCLOSE", _
		"DTCOUPON", _
		"DTCREATED", _
		"DTDUE", _
		"DTEFF", _
		"DTEND", _
		"DTEXPIRE", _
		"DTMAT", _
		"DTNEXT", _
		"DTOPEN", _
		"DTPMTDUE", _
		"DTPMTPRC", _
		"DTPOSTED", _
		"DTPOSTEND", _
		"DTPOSTSTART", _
		"DTPROFUP", _
		"DTPURCHASE", _
		"DTSEEN", _
		"DTSERVER", _
		"DTSETTLE", _
		"DTSTART", _
		"DTTRADE", _
		"DTUPDATE", _
		"DTUSER", _
		"DTXFERPRC", _
		"DTXFERPRJ", _
		"DTYIELDASOF", _
		"DURATION", _
		"EMAIL", _
		"EVEPHONE", _
		"EXTDPMTCHK", _
		"EXTDPMTCHK2", _
		"EXTDPMTDSC2", _
		"EXTDPMTFOR", _
		"FAXPHONE", _
		"FEE", _
		"FEEMSG", _
		"FEES", _
		"FIASSETCLASS", _
		"FICERTID", _
		"FIID", _
		"FINALAMT", _
		"FINAME", _
		"FINCHG", _
		"FIRSTNAME", _
		"FITID", _
		"FRACCASH", _
		"FREQ", _
		"FROM", _
		"GAIN", _
		"GENUSERKEY", _
		"GETMIMESUP", _
		"GROUPID", _
		"HASEXTDPMT", _
		"HELDINACCT", _
		"HELPMESSAGE", _
		"IDSCOPE", _
		"IMAGEURL", _
		"INCBAL", _
		"INCIMAGES", _
		"INCLUDE", _
		"INCLUDEBILLPMTSTATUS", _
		"INCLUDEBILLSTATUS", _
		"INCLUDECOUNTS", _
		"INCLUDEDETAIL", _
		"INCLUDESTATUSHIST", _
		"INCLUDESUMMARY", _
		"INCOMETYPE", _
		"INCOO", _
		"INITIALAMT", _
		"INTLXFERFEE", _
		"INTU.BID", _
		"INVACCTTYPE", _
		"INVDATE", _
		"INVDESC", _
		"INVNO", _
		"INVPAIDAMT", _
		"INVTOTALAMT", _
		"ITA.CAUSALE", _
		"LANGUAGE", _
		"LASTNAME", _
		"LIMITPRICE", _
		"LITMAMT", _
		"LITMCODE", _
		"LITMDESC", _
		"LOAD", _
		"LOGO", _
		"LOSTSYNC", _
		"MAILSUP", _
		"MARGINBALANCE", _
		"MARKDOWN", _
		"MARKUP", _
		"MAX", _
		"MEMO", _
		"MEMO2", _
		"MESSAGE", _
		"MESSAGE2", _
		"MFTYPE", _
		"MIDDLENAME", _
		"MIN", _
		"MINPMTDUE", _
		"MINUNITS", _
		"MKTGINFO", _
		"MKTVAL", _
		"MODELWND", _
		"MODPENDING", _
		"N", _
		"NAME", _
		"NAMEACCTHELD", _
		"NEEDTANPAYEE", _
		"NEEDTANPMT", _
		"NEEDTANTRANSFER", _
		"NEWUNITS", _
		"NEWUSERPASS", _
		"NINSTS", _
		"NONCE", _
		"NOTIFYDESIRED", _
		"NOTIFYWILLING", _
		"NUMERATOR", _
		"OFXSEC", _
		"OLDUNITS", _
		"ONETIMEPASS", _
		"OODNLD", _
		"OPTACTION", _
		"OPTBUYTYPE", _
		"OPTIONLEVEL", _
		"OPTSELLTYPE", _
		"OPTTYPE", _
		"PARVALUE", _
		"PAYACCT", _
		"PAYANDCREDIT", _
		"PAYEEID", _
		"PAYEEID2", _
		"PAYEELSTID", _
		"PAYEELSTID2", _
		"PAYEEMODPENDNG", _
		"PAYINSTRUCT", _
		"PERCENT", _
		"PHONE", _
		"PINCH", _
		"PMTBYADDR", _
		"PMTBYPAYEEID", _
		"PMTBYXFER", _
		"PMTFOR", _
		"PMTINSTRUMENTTYPE", _
		"PMTTYPE", _
		"POSDNLD", _
		"POSTALCODE", _
		"POSTPRCWND", _
		"POSTYPE", _
		"PREAUTH", _
		"PREAUTHTOKEN", _
		"PREFETCHURL", _
		"PREVBAL", _
		"PROCDAYSOFF", _
		"PROCENDTM", _
		"PURANDADV", _
		"PWTYPE", _
		"RATING", _
		"REASON", _
		"RECSRVRTID", _
		"RECSRVRTID2", _
		"REFNUM", _
		"REFRESH", _
		"REFRESHSUPT", _
		"REINVCG", _
		"REINVDIV", _
		"REJECTIFMISSING", _
		"RELFITID", _
		"RELTYPE", _
		"RESPFILEER", _
		"RESTRICT", _
		"RESTRICTION", _
		"REVERSALFEES", _
		"REVERSALFITID", _
		"SECLISTRQDNLD", _
		"SECNAME", _
		"SECURED", _
		"SECURITYNAME", _
		"SELLALL", _
		"SELLREASON", _
		"SELLTYPE", _
		"SESSCOOKIE", _
		"SEVERITY", _
		"SHORTBALANCE", _
		"SHPERCTRCT", _
		"SIC", _
		"SIGNONREALM", _
		"SPACES", _
		"SPECIAL", _
		"SPNAME", _
		"SRVRTID", _
		"SRVRTID2", _
		"STATE", _
		"STATUSMODBY", _
		"STOCKTYPE", _
		"STOPPRICE", _
		"STPCHKFEE", _
		"STRIKEPRICE", _
		"STSVIAMODS", _
		"SUBACCT", _
		"SUBACCTFROM", _
		"SUBACCTFUND", _
		"SUBACCTSEC", _
		"SUBACCTTO", _
		"SUBJECT", _
		"SUPPORTDTAVAIL", _
		"SUPTXDL", _
		"SVC", _
		"SVC2", _
		"SVCSTATUS", _
		"SVCSTATUS2", _
		"SWITCHALL", _
		"SYNCERROR", _
		"TABLENAME", _
		"TAN", _
		"TAXES", _
		"TAXEXEMPT", _
		"TAXID", _
		"TEMPPASS", _
		"TFERACTION", _
		"TICKER", _
		"TO", _
		"TOKEN", _
		"TOKEN2", _
		"TOKENONLY", _
		"TOTAL", _
		"TOTALFEES", _
		"TOTALINT", _
		"TRANDNLD", _
		"TRANSPSEC", _
		"TRNAMT", _
		"TRNTYPE", _
		"TRNUID", _
		"TSKEYEXPIRE", _
		"TSPHONE", _
		"TYPEDESC", _
		"UNIQUEID", _
		"UNIQUEIDTYPE", _
		"UNITPRICE", _
		"UNITS", _
		"UNITSSTREET", _
		"UNITSUSER", _
		"UNITTYPE", _
		"URL", _
		"URL2", _
		"URLGETREDIRECT", _
		"USEHTML", _
		"USERID", _
		"USERKEY", _
		"USERPASS", _
		"USPRODUCTTYPE", _
		"VALIDATE", _
		"VALUE", _
		"VER", _
		"WITHHOLDING", _
		"XFERDAYSWITH", _
		"XFERDEST", _
		"XFERDFLTDAYSTOPAY", _
		"XFERPRCCODE", _
		"XFERSRC", _
		"YIELD", _
		"YIELDTOCALL", _
		"YIELDTOMAT" _
	)

	Set ElementsDict = CreateObject("Scripting.Dictionary")
	Dim e
	For Each e In aElements
		ElementsDict.Add e,""
	Next
End Sub

' checks to see if the given tag is an OFX Element (see above)
Function IsOFXElement(sTag)
	IsOFXElement = ElementsDict.Exists(sTag)
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
' OFX is almost there, but "Elements" in OFX-speak do not have a closing tag,
' which we have to add. "Aggregates" and "Records" do have a closing tag, so we
' don't add another one!
' also each line seems to start (apart from white space) with a tag!
' I hope this heuristic holds good, otherwise the routine below will need rewriting.
Function OFXtoXML()
	Dim sLine
	Dim iGT
	Dim sTag
	Dim sText
	sXML = ""
	Rewind
	sLine = Trim(Replace(ReadLine(),vbTab, ""))
	Rewind
	If StartsWith(sLine, "<?xml") Then
		Do While Not AtEOF()
			sLine = Trim(ReadLine())
			sXML = sXML & sLine & vbcrlf
		Loop
	Else
		Do While Not AtEOF()
			sLine = Trim(Replace(ReadLine(),vbTab, ""))
			' AF-> Search OFX tag
			If StartsWith(sLine, "<OFX") Then
				sXML = sLine & vbcrlf
				Exit Do
			End If
		Loop
		Do While Not AtEOF()
			sLine = Trim(Replace(ReadLine(),vbTab, ""))
			If Left(sLine, 1) = "<" Then
				If Left(sLine, 2) <> "</" Then
					iGT = InStr(sLine, ">")
					If iGT > 0 Then
						sTag = Mid(sLine, 2, iGT-2)
						If IsOFXElement(sTag) Then
							sText = FixXMLEncoding(Mid(sLine, iGT+1))
							sLine = "<" & sTag & ">" & sText & "</" & sTag & ">"
						End If
					End If
				End If
			End If
			sXML = sXML & sLine & vbcrlf
		Loop
	End If
	OFXtoXML = sXML
	' MsgBox sXML, vbOkOnly, "OFX Processing Debug sXML"
End Function

Function LoadAsXML()
	Dim bRet: bRet = False
	Dim sXML
	sXML = OFXtoXML()
	If xDoc Is Nothing Then
		Set xDoc = CreateObject("MSXML2.DOMDocument")
		xDoc.async = False
		xDoc.validateOnParse = False
	End If
	bRet = xDoc.loadXML(sXML)
	If bRet Then
		bRet=(xDoc.documentElement.NodeName = "OFX")
		If Not bRet Then
			MsgBox "Document is not an OFX file!",vbokonly+vbcritical,"OFX Processing Error"
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
	If InStr(sAmt, ",") = 0 Then
		ParseAmount = ParseNumber(sAmt, ".")
	Else
		ParseAmount = ParseNumber(sAmt, ",")
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
	Dim bRet: bRet = False
	Dim i
	Dim j
	Dim sTmp

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

	Session.ServerTime = ParseDate(NodeText(xDoc.selectSingleNode("/OFX/SIGNONMSGSRSV1/SONRS/DTSERVER")))
	Bcfg.IntuitBankId = NodeText(xDoc.selectSingleNode("/OFX/SIGNONMSGSRSV1/SONRS/INTU.BID"))

	Set xStmts = xDoc.selectNodes("/OFX/BANKMSGSRSV1/STMTTRNRS/STMTRS")
'	MsgBox "Statements found: " & xStmts.length
	For i=0 To xStmts.length-1
		Set xStmt = xStmts.item(i)
' process statement here!
		Set Stmt = NewStatement()
		Set xAcct = xStmt.selectSingleNode("BANKACCTFROM")
		Stmt.BankName = NodeText(xAcct.selectSingleNode("BANKID"))
		Stmt.Acct = NodeText(xAcct.selectSingleNode("ACCTID"))
		Stmt.ClosingBalance.Amt = ParseAmount(NodeText(xStmt.selectSingleNode("LEDGERBAL/BALAMT")))
		Stmt.ClosingBalance.Ccy = NodeText(xStmt.selectSingleNode("CURDEF"))
		Stmt.ClosingBalance.BalDate = ParseDate(NodeText(xStmt.selectSingleNode("LEDGERBAL/DTASOF")))
		Stmt.OpeningBalance.Ccy = NodeText(xStmt.selectSingleNode("CURDEF"))
		Stmt.OpeningBalance.BalDate = ParseDate(NodeText(xStmt.selectSingleNode("BANKTRANLIST/DTSTART")))
		Set xTxns = xStmt.selectNodes("BANKTRANLIST/STMTTRN")
'		MsgBox "Transactions found: " & xTxns.length
		For j=0 To xTxns.length-1
			Set xTxn = xTxns.item(j)
' process transaction here!
			NewTransaction
			Txn.Amt = ParseAmount(NodeText(xTxn.selectSingleNode("TRNAMT")))
			Txn.BookDate = ParseDate(NodeText(xTxn.selectSingleNode("DTPOSTED")))
			sTmp = NodeText(xTxn.selectSingleNode("DTAVAIL"))
			If Len(sTmp) > 0 Then
				Txn.ValueDate = ParseDate(sTmp)
'			Else
'				Txn.ValueDate = Txn.BookDate
			End If
			Txn.Memo = NodeText(xTxn.selectSingleNode("MEMO"))
'			MsgBox Txn.Memo
			Txn.Payee = NodeText(xTxn.selectSingleNode("NAME"))
			If Len(Txn.Payee) = 0 Then
				Txn.Payee = NodeText(xTxn.selectSingleNode("PAYEE/NAME"))
			End if
			Txn.TxnType = NodeText(xTxn.selectSingleNode("TRNTYPE"))
			Txn.FITID = NodeText(xTxn.selectSingleNode("FITID"))
			Txn.CheckNum = NodeText(xTxn.selectSingleNode("CHECKNUM"))
		Next
	Next
	LoadTextFile = bRet
End Function
