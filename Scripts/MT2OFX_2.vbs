Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/MT2OFX.vbs 35    24/11/09 22:05 Colin $"

' MT2OFX.VBS
' This file contains constants and global functions for use in the
' scripting system of MT2OFX.

' Structured :86: subfields
Public Const sfBookingText = 0
Public Const sfBatchNum = 10
Public Const sfDetails0 = 20
Public Const sfDetailsLast = 29
Public Const sfBankPayee = 30
Public Const sfAcctPayee = 31
Public Const sfNamePayee = 32
Public Const sfNamePayee2 = 33
Public Const sfTextCodeSupplement = 34
Public Const sfExtra0 = 60
Public Const sfExtraLast = 63

' Structured :NS: subfields
Public Const nsfApplication0 = 1
Public Const nsfApplicationLast = 14
Public Const nsfPayee = 15
Public Const nsfPayee2 = 16
Public Const nsfBookingText = 17
Public Const nsfJournalNumber = 18
Public Const nsfBookingTime = 19	' format: HHMM
Public Const nsfNumTransactions = 20
Public Const nsfAccountHolder = 22
Public Const nsfAccountDescription = 23
Public Const nsfInterestRate = 24
Public Const nsfDuration = 25		' format: DDMMYYDDMMYY (from+to)
Public Const nsfTargetAccount = 26
Public Const nsfCounterValue = 27
Public Const nsfOriginalValue = 28
Public Const nsfExchangeRate = 29
Public Const nsfBankSortCode = 30
Public Const nsfParticipantNumber = 31
Public Const nsfAccountType = 32
Public Const nsfPayeeBank = 33
Public Const nsfPayeeAccount = 34

' Script property data types
Public Const ptString = 1
Public Const ptInteger = 2
Public Const ptDate = 3
Public Const ptBoolean = 4
Public Const ptFloat = 5
Public Const ptChoice = 6
Public Const ptCurrency = 7

' CSV Field Codes
Public Const fldSkip = 0
Public Const fldAccountNum = 1
Public Const fldCurrency = 2
Public Const fldClosingBal = 3
Public Const fldAvailBal = 4
Public Const fldBookDate = 5
Public Const fldValueDate = 6
Public Const fldAmtCredit = 7
Public Const fldAmtDebit = 8
Public Const fldMemo = 9
Public Const fldBalanceDate = 10
Public Const fldAmount = 11
Public Const fldPayee = 12
Public Const fldTransactionDate = 13
Public Const fldTransactionTime = 14
Public Const fldChequeNum = 15
Public Const fldCheckNum = 15
Public Const fldFITID = 16
Public Const fldEmpty = 17	' field is ignored but MUST be empty for recognition
Public Const fldBranch = 18
Public Const fldSign = 19	' + or -
Public Const fldCategory = 20
Public Const fldPayeeCity = 21
Public Const fldPayeeState = 22
Public Const fldPayeeZip = 23
Public Const fldPayeeCountry = 24
Public Const fldPayeePhone = 25
Public Const fldPayeeAddress1 = 26
Public Const fldPayeeAddress2 = 27
Public Const fldPayeeAddress3 = 28
Public Const fldPayeeAddress4 = 29
Public Const fldPayeeAddress5 = 30
Public Const fldPayeeAddress6 = 31
Public Const fldPayeeAcctNum = 32
Public Const fldPayeeAcctBank = 33
Public Const fldPayeeAcctBranch = 34
Public Const fldPayeeAcctType = 35
Public Const fldPayeeAcctKey = 36
' 20090827 CS: add five user fields which can be handled by callbacks
Public Const fldUser1 = 101
Public Const fldUser2 = 102
Public Const fldUser3 = 103
Public Const fldUser4 = 104
Public Const fldUser5 = 105

' A few useful codepages
Public Const CP_ACP = 0
Public Const CP_OEMCP = 1
Public Const CP_UTF16 = 1200
Public Const CP_UTF7 = 65000
Public Const CP_UTF8 = 65001

' 20090904 CS: Support for continuation lines
Public Const txnlineNORMAL = -1        ' = CInt(True)
Public Const txnlineSKIP = 0           ' = CInt(False)
Public Const txnlineCONTINUE = 1       ' continue previous transaction
Public Const txnlineNOTRANSACTION = 2  ' process, but does not need a transaction
Public Const txnlineNORMALNEWSTATEMENT = 3         ' process as normal, start a new statement first
Public Const txnlineNOTRANSACTIONNEWSTATEMENT = 4  ' no transaction, but start a new statement anyway

Public NODATE
Public ScriptName

Sub Initialise()
' ensure locale is initialised properly!
	SetLocale(0)
	NODATE = DateSerial(0,0,0)
End Sub

' 20071102 CS: VersionAtLeast added so we can check for codepage support in V3.5
' sMinVer: String, "x" or "x.y" or "x.y.z" where x,y,z are major,minor,build numbers from the program version
Function VersionAtLeast(sMinVer)
	Dim iMajor, iMinor, iBuild
	iMajor = 0: iMinor = 0: iBuild = 0
	Dim vParts
	vParts = Split(sMinVer, ".")
	iMajor = CLng(vParts(0))
	If UBound(vParts) > 0 Then
		iMinor = CLng(vParts(1))
		If UBound(vParts) > 1 Then
			iBuild = CLng(vParts(2))
		End If
	End If
	vParts = Split(Version, ".")
	VersionAtLeast = False
	If CLng(vParts(0)) < iMajor Then
		Exit Function
	ElseIf CLng(vParts(0)) = iMajor Then
		If CLng(vParts(1)) < iMinor Then
			Exit Function
		ElseIf CLng(vParts(1)) = iMinor Then
			If CLng(vParts(2)) < iBuild Then
				Exit Function
			End If
		End If
	End If
	VersionAtLeast = True
End Function

Function CheckVersion
	CheckVersion = True
	If ScriptEngineMajorVersion < 5 Then
		MsgBox "A newer version of Microsoft Scripting is required than """ _
			& GetScriptEngineInfo() & """ currently installed." _ 
			& vbCrLf _
			& "Please upgrade.", vbOkOnly+vbCritical, "MT2OFX Scripting"
		CheckVersion = False
	End If
End Function

Function GetScriptEngineInfo
   Dim s
   s = ""   ' Build string with necessary info.
   s = ScriptEngine & " Version "
   s = s & ScriptEngineMajorVersion & "."
   s = s & ScriptEngineMinorVersion & "."
   s = s & ScriptEngineBuildVersion 
   GetScriptEngineInfo = s   ' Return the results.
End Function

Sub DumpObjects(sFile)
	Dim fso, i, x, j, k, s
	Dim MyFile
	Set fso = CreateObject("Scripting.FileSystemObject")
	If Len(sFile) = 0 Then
		Exit Sub
	End If
	Set MyFile = fso.CreateTextFile(sFile, True)


	MyFile.WriteLine("Script Engine: " & GetScriptEngineInfo())
	MyFile.WriteLine("")
	
	MyFile.WriteLine("Session")
	MyFile.WriteLine("=======")
	MyFile.WriteLine("Program Version  : " & Version)
	MyFile.WriteLine("FileIn           : " & Session.FileIn)
	MyFile.WriteLine("FileOut          : " & Session.FileOut)
	MyFile.WriteLine("BankID           : " & Session.BankID)
	MyFile.WriteLine("OutputFileType   : " & Session.OutputFileType)
	MyFile.WriteLine("PayeeMapFile     : " & Session.PayeeMapFile)
	MyFile.WriteLine("PayeeMapIgnoreCase: " & Session.PayeeMapIgnoreCase)
	
	MyFile.WriteLine("")

	MyFile.WriteLine("Program Configuration")
	MyFile.WriteLine("=====================")
	MyFile.WriteLine("MemoDelimiter           : " & Cfg.MemoDelimiter)
	MyFile.WriteLine("OutputFileType          : " & Cfg.OutputFileType)
	MyFile.WriteLine("HomePage                : " & Cfg.HomePage)
	MyFile.WriteLine("PayeeCase               : " & Cfg.PayeeCase)
	MyFile.WriteLine("AutoBrowse              : " & Cfg.AutoBrowse)
' 20050421 CS: I thought this was already removed!!!!
'	MyFile.WriteLine("AutoStartImport         : " & Cfg.AutoStartImport)
	MyFile.WriteLine("PromptForOutput         : " & Cfg.PromptForOutput)
	MyFile.WriteLine("BookDateMode            : " & Cfg.BookDateMode)
	MyFile.WriteLine("CompressSpaces          : " & Cfg.CompressSpaces)
	MyFile.WriteLine("LogFile                 : " & Cfg.LogFile)
	MyFile.WriteLine("TxnDumpFile             : " & Cfg.TxnDumpFile)
	MyFile.WriteLine("ScriptDebugLevel        : " & Cfg.ScriptDebugLevel)
	MyFile.WriteLine("ScriptTimeout           : " & Cfg.ScriptTimeout)
	MyFile.WriteLine("ScriptEditor            : " & Cfg.ScriptEditor)
' AutoText
' AutoMT940
' PayeeMapFile
' PayeeMapIgnoreCase
	MyFile.WriteLine("")

	MyFile.WriteLine("Bank Configuration")
	MyFile.WriteLine("==================")
	MyFile.WriteLine("IDString           : " & Bcfg.IDString)
	MyFile.WriteLine("Structured86       : " & BCfg.Structured86)
	MyFile.WriteLine("SkipEmptyMemoFields: " & BCfg.SkipEmptyMemoFields)
	MyFile.WriteLine("FindPayeeFn        : " & BCfg.FindPayeeFn)
	MyFile.WriteLine("FindTxnDateFn      : " & BCfg.FindTxnDateFn)
	MyFile.WriteLine("FindServerTimeFn   : " & BCfg.FindServerTimeFn)
	MyFile.WriteLine("ScriptFile         : " & BCfg.ScriptFile)
	MyFile.WriteLine("")

	For k=1 to Session.Statements.Count
		Set s=Session.Statements(k)
		MyFile.WriteLine("Statement")
		MyFile.WriteLine("=========")
		MyFile.WriteLine("Acct        :" & s.Acct)
		MyFile.WriteLine("BankName    :" & s.BankName)
		MyFile.WriteLine("StatementID :" & s.StatementID)
		MyFile.WriteLine("OpeningBalance :" & _
			s.OpeningBalance.Amt & s.OpeningBalance.Ccy & " @ " & s.OpeningBalance.BalDate)
		MyFile.WriteLine("ClosingBalance :" & _
			s.ClosingBalance.Amt & s.ClosingBalance.Ccy & " @ " & s.ClosingBalance.BalDate)
		MyFile.WriteLine("AvailableBalance :" & _
			s.AvailableBalance.Amt & s.AvailableBalance.Ccy & " @ " & s.AvailableBalance.BalDate)
		MyFile.WriteLine("OFXStatementID :" & s.OFXStatementID)
		MyFile.WriteLine("")
		If s.NonSwift.Fields.Count > 0 Then
			MyFile.WriteLine("NS-Records (" & s.NonSwift.Fields.Count & ")")
			MyFile.WriteLine("==========")
			For i=1 To s.NonSwift.Fields.Count
				MyFile.WriteLine(s.NonSwift.Fields(i).FieldNum & ": " & s.NonSwift.Fields(i).Text)
			Next
			MyFile.WriteLine("")
		End If

		MyFile.WriteLine("Transactions (" & s.Txns.Count & ")")
		MyFile.WriteLine("============")
		For i=1 To s.Txns.Count
			Set x = s.Txns(i)
			MyFile.WriteLine("")
			MyFile.WriteLine("Transaction " & i & " of " & s.Txns.Count)
			MyFile.WriteLine("SkipPayeeMap : " & x.SkipPayeeMapping)
			MyFile.WriteLine("ClearedStatus: " & x.ClearedStatus)
			MyFile.WriteLine("Amt          : " & x.Amt)
			MyFile.WriteLine("BookDate     : " & x.BookDate)
			MyFile.WriteLine("ValueDate    : " & x.ValueDate)
			MyFile.WriteLine("IsReversal   : " & x.IsReversal)
			MyFile.WriteLine("BookingCode  : " & x.BookingCode)
			MyFile.WriteLine("Reference    : " & x.Reference)
			MyFile.WriteLine("BankReference: " & x.BankReference)
			MyFile.WriteLine("Memo         : " & x.Memo)
			If Bcfg.Structured86 Then
				MyFile.WriteLine("Structured :86: subfields")
				MyFile.WriteLine("=========================")
				MyFile.WriteLine("BusCode  : " & x.Str86.BusCode)
				For j=1 To x.Str86.Fields.Count
					MyFile.WriteLine(x.Str86.Fields(j).FieldNum & ": " & x.Str86.Fields(j).Text)
				Next
			End If
' 20051126 CS Updated for full Payee model
			If x.Payee.IsSimple Then
				MyFile.WriteLine("Payee        : " & x.Payee)
			Else
				If Len(x.Payee.Name) > 0 Then
					MyFile.WriteLine("Payee        : " & x.Payee.Name)
				End If
				For j=1 To x.Payee.LastUsedAddressLine
					If Len(x.Payee.Addr(j)) > 0 Then
						MyFile.WriteLine("Addr " & CStr(j) & "       : " & x.Payee.Addr(j))
					End If
				Next
				If Len(x.Payee.City) > 0 Then
					MyFile.WriteLine("City         : " & x.Payee.City)
				End If
				If Len(x.Payee.State) > 0 Then
					MyFile.WriteLine("State        : " & x.Payee.State)
				End If
				If Len(x.Payee.PostalCode) > 0 Then
					MyFile.WriteLine("Postal Code  : " & x.Payee.PostalCode)
				End If
				If Len(x.Payee.Country) > 0 Then
					MyFile.WriteLine("Country      : " & x.Payee.Country)
				End If
				If Len(x.Payee.Phone) > 0 Then
					MyFile.WriteLine("Phone        : " & x.Payee.Phone)
				End If
			End If
			If Len(x.Category) > 0 Then
				MyFile.WriteLine("Category     : " & x.Category)
			End If
			MyFile.WriteLine("FITID        : " & x.FITID)
			MyFile.WriteLine("CheckNum     : " & x.CheckNum)
			MyFile.WriteLine("TxnType      : " & x.TxnType)
			If x.TxnDateValid Then MyFile.WriteLine("TxnDate      : " & x.TxnDate)
			If x.NonSwift.Fields.Count > 0 Then
				MyFile.WriteLine("NS-Records (" & x.NonSwift.Fields.Count & ")")
				MyFile.WriteLine("==========")
				For j=1 To x.NonSwift.Fields.Count
					MyFile.WriteLine(x.NonSwift.Fields(j).FieldNum & ": " & x.NonSwift.Fields(j).Text)
				Next
			End If
' 20041129 CS Added Split handling
			If x.Splits.Count > 0 Then
				MyFile.WriteLine("Splits (" & x.Splits.Count & ")")
				MyFile.WriteLine("======")
				For j=1 To x.Splits.Count
					MyFile.WriteLine(" Amount   : " & x.Splits(j).Amt)
					MyFile.WriteLine(" Category : " & x.Splits(j).Category)
					MyFile.WriteLine(" Memo     : " & x.Splits(j).Memo)
				Next
			End If
		Next	' transaction
		MyFile.WriteLine("")
	Next	' statement
End Sub

Sub LogProgress(sBank, sProcedure)
	If Cfg.ScriptDebugLevel > 5 Then
		MsgBox "In " & sProcedure & " (" & sBank & ") "
	End If
End Sub

Function StartsWith(s, Prefix)
	StartsWith = (Left(s,Len(Prefix)) = Prefix)
End Function

Dim LastMemo	' last non-blank memo field seen
Sub ConcatMemo(s)
	If s = "" Then
		Exit Sub
	End If
	If Len(Txn.Memo) > 0 Then
		Txn.Memo = Txn.Memo & Cfg.MemoDelimiter
	End If
	Txn.Memo = Txn.Memo & s
	LastMemo = s
End Sub

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

Dim MonthDict
Set MonthDict = CreateObject("Scripting.Dictionary")

Sub InitialiseMonths(aMonths)
	Dim i
	If IsEmpty(aMonths) Then
		For i=1 To 12
			MonthDict(lcase(MonthName(i, False))) = i
		Next
		For i=1 To 12
			MonthDict(lcase(MonthName(i, True))) = i
		Next
	Else
		For i=0 to UBound(aMonths)
			MonthDict(lcase(aMonths(i))) = (i Mod 12) + 1
		Next
	End If
End Sub

Dim ParseDateError

Function ParseDateEx(sDate, mySeq, mySep)
	Dim iYear, iMonth, iDay
	Dim iPart1, iPart2, iPart3
	Dim aSplit
	Dim i
	Dim re, m, sTmp
	Dim bNoYear
	ParseDateEx = NODATE
	ParseDateError = ""
	bNoYear = False
	If Len(sDate) = 0 Then
		Exit Function
	End If
	If mySep = "" Then
		If Not IsNumeric(sDate) Then
			ParseDateError = "Date  '" & sDate & "' is not numeric and no separator is defined."
			Exit Function
		End If
		If Len(sDate) = 6 Then
			iPart1 = Left(sDate, 2)
			iPart2 = Mid(sDate, 3, 2)
			iPart3 = Mid(sDate, 5, 2)
		Elseif Len(sDate) = 8 Then
			If mySeq = "YMD" Then
				iPart1 = Left(sDate, 4)
				iPart2 = Mid(sDate, 5, 2)
				iPart3 = Mid(sDate, 7, 2)
			Else
				iPart1 = Left(sDate, 2)
				iPart2 = Mid(sDate, 3, 2)
				iPart3 = Mid(sDate, 5, 4)
			End If
		ElseIf Len(sDate) = 4 Then
			iPart1 = Left(sDate, 2)
			iPart2 = Mid(sDate, 2)
			iPart3 = "0"
			bNoYear = True
		Else
			ParseDateError = "Don't know what to do with date of length " & Len(sDate)
			Exit Function
		End If
	Else
		If Len(mySep) = 1 Then	' single separator (legacy case)
			aSplit = Split(sDate, mySep)
			Select Case UBound(aSplit)
			Case 1
				iPart1 = Trim(aSplit(0))
				iPart2 = Trim(aSplit(1))
				iPart3 = "0"
				bNoYear = True
			Case 2
				iPart1 = Trim(aSplit(0))
				iPart2 = Trim(aSplit(1))
				iPart3 = Trim(aSplit(2))
			Case Else
				ParseDateError = "Unable to parse date '" & sDate & "' - expecting separator '" & mySep & "'."
				Exit Function
			End Select
		Else ' multi-char separator - use RegExp to split the date
			Set re = New RegExp
			re.Global = False
			sTmp = Replace(mySep, ".", "\.")
			re.Pattern = "([^" & sTmp & "]+)[" & sTmp & "]+([^" & sTmp & "]+)(?:[" & sTmp & "]+([^" & sTmp & "]+))?"
			re.IgnoreCase = False
			Set m = re.Execute(sDate)
			If m.Count < 1 Then
				ParseDateError = "Date '" & sDate & "' failed to match pattern '" & re.Pattern & "'"
				Exit Function
			Else
				With m(0)
					iPart1 = .SubMatches(0)
					iPart2 = .SubMatches(1)
					If .SubMatches.Count = 3 Then
						iPart3 = .SubMatches(2)
					Else
						iPart3 = "0"
						bNoYear = True
					End If
				End With
			End If
		End If
' 20050608 CS: lose any time part after a Space
		i = InStr(iPart3, " ")
		If i>0 Then
			iPart3 = Left(iPart3, i-1)
		End If
	End If
	Select Case mySeq
	Case "MDY"
		iMonth = iPart1
		iDay = iPart2
		iYear = iPart3
	Case "DMY"
		iDay = iPart1
		iMonth = iPart2
		iYear = iPart3
	Case "YMD"
		iYear = iPart1	
		iMonth = iPart2
		iDay = iPart3
	Case Else
		ParseDateError = "Unknown date sequence '" & mySeq & "'"
		Exit Function
	End Select
	If Not IsNumeric(iMonth) Then
		If MonthDict.Exists(lcase(iMonth)) Then
			iMonth = MonthDict(lcase(iMonth))
		Else
			ParseDateError = "Unrecognised month '" & iMonth & "'"
			Exit Function
		End If
	End If
	If Not IsNumeric(iDay) Or _
		Not IsNumeric(iMonth) Or _
		Not IsNumeric(iYear) Then
		ParseDateError = "Unable to parse date '" & sDate & "' - non-numeric field found."
		Exit Function
	End If
	If iYear < 100 Then
		If iYear > 70 Then
			iYear = iYear + 1900
		Else
			iYear = iYear + 2000
		End If
	End If
	If Not (iDay < 1 Or iDay>31 Or iMonth < 1 Or iMonth>12 Or iYear < 1970 Or iYear > 2100) Then
		ParseDateEx = DateSerial(iYear, iMonth, iDay)
      If ParseDateEx = NODATE Then
        ParseDateError = "Error from DateSerial(" & iYear & "," & iMonth & "," & iDay & ")"
      Else
        If bNoYear Then
            ParseDateEx = MostRecent(ParseDateEx)
        End If
      End If
   Else
      ParseDateError = "Improper date: d=" & iDay & ", m=" & iMonth & ", y=" & iYear
	End If
	
End Function

' Most recent: where d must be the same date or earlier
Function MostRecentEx(d, dNow)
	Dim dNew
	dNew = DateSerial(Year(dNow), Month(d), Day(d))
	If dNew > dNow Then
		dNew = DateAdd("yyyy", -1, dNew)
	End If
	MostRecentEx = dNew
End Function

Function MostRecent(d)
	MostRecent = MostRecentEx(d, Now())
End Function

' Closest: fix the year of d such that it is closest to dNow (which must contain a real year)
Function ClosestEx(d, dNow)
	Dim dNew
	dNew = DateSerial(Year(dNow), Month(d), Day(d))
	If DateAdd("m", -6, dNew) > dNow Then
		dNew = DateAdd("yyyy", -1, dNew)
	End If
	ClosestEx = dNew	
End Function

Function Closest(d)
	Closest = ClosestEx(d, Now())
End Function

Function StringMatches(s, pat)
	Dim r
	Set r=New RegExp
	r.Global = False
	r.Pattern = pat
	r.IgnoreCase = False
	StringMatches = r.Test(s)
End Function

Public Function ExtractDigits(sString)
	Dim s, c
	Dim i
	For i=1 To Len(sString)
		c = Mid(sString, i, 1)
		If c>="0" And c<= "9" Then
			s = s & c
		End If
	Next
	ExtractDigits = s
End Function

' 20090827: Auto-detect CSV separator
Public Function FindCSVSeparator(sLine)
    Dim iCount, vFields, sChar

    vFields = ParseLineDelimited(sLine, ",")
    iCount = UBound(vFields)
    sChar = ","

    vFields = ParseLineDelimited(sLine, ";")
    If UBound(vFields) > iCount Then
        iCount = UBound(vFields)
        sChar = ";"
    ElseIf UBound(vFields) = iCount Then
        If InStr(sLine, ",") > InStr(sLine, ";") Then
            sChar = ";"
        End If
    End If

    vFields = ParseLineDelimited(sLine, vbTab)
    If UBound(vFields) > iCount Then
        iCount = UBound(vFields)
        sChar = vbTab
    End If
    
    If iCount >= 3 Then
        FindCSVSeparator = sChar
    Else
        FindCSVSeparator = ""
    End If
End Function

' Mastercard: Must have a prefix of 51 to 55, and must be 16 digits in length.
' Visa: Must have a prefix of 4, and must be either 13 or 16 digits in length.
' American Express: Must have a prefix of 34 or 37, and must be 15 digits in length.
' Diners Club: Must have a prefix of 300 to 305, 36, or 38, and must be 14 digits in length.
' Discover: Must have a prefix of 6011, and must be 16 digits in length.
' JCB: Must have a prefix of 3, 1800, or 2131, and must be either 15 or 16 digits in length.
Public Const ccInvalid = -1
Public Const ccUnknown = 0
Public Const ccMastercard = 1
Public Const ccVisa = 2
Public Const ccAmex = 3
Public Const ccDiners = 4
Public Const ccDiscover = 5
Public Const ccJCB = 6

Public Function ValidateCreditCard(sCard)
	Dim NumberLength, iPrefix, ShouldLength
	
	ValidateCreditCard = ccInvalid
	If Len(sCard) < 10 Then
		Exit Function
	End If
	If Not IsNumeric(sCard) Then
		Exit Function
	End If
	If Not LuhnCheck(sCard) Then
		Exit Function
	End If

    '2) Do the first four digits fit within proper ranges?
    '     If so, who's the card issuer and how long should the number be?
    ValidateCreditCard = ccUnknown
    NumberLength = Len(sCard)
    iPrefix = CInt(Left(sCard, 4))
    If (iPrefix>=3000 And iPrefix<=3059) Or (iPrefix>=3600 And iPrefix<=3699) Or (iPrefix>=3800 And iPrefix<=3889) Then
		ValidateCreditCard = ccDiners
        ShouldLength = 14
    ElseIf (iPrefix>=3400 And iPrefix<=3499) Or (iPrefix>=3700 And iPrefix<=3799) Then
        ValidateCreditCard = ccAmex
        ShouldLength = 15
    ElseIf (iPrefix>=3528 And iPrefix<=3589) Then
        ValidateCreditCard = ccJCB
        ShouldLength = 16
    ElseIf (iPrefix>=4000 And iPrefix<=4999) Then
        ValidateCreditCard = ccVisa
        If NumberLength > 14 Then
            ShouldLength = 16
        ElseIf NumberLength < 14 Then
            ShouldLength = 13
        Else
         	ValidateCreditCard = ccInvalid
            Exit Function
        End If
	ElseIf (iPrefix>=5100 And iPrefix<=5599) Then
        ValidateCreditCard = ccMastercard
        ShouldLength = 16
    ElseIf iPrefix=6011 Or (iPrefix>=6500 And iPrefix<=6509) Then
    	ValidateCreditCard = ccDiscover
        ShouldLength = 16
    Else
    	ValidateCreditCard = ccUnknown
        Exit Function
    End If

    '3) Is the number the right length?
    If NumberLength <> ShouldLength Then
    	ValidateCreditCard = ccInvalid
        Exit Function
    End If
   
End Function

Public Function LuhnCheck(sCard)
    'This works for numbers up to 255 characters long.
    'For longer numbers, increase variable data types as needed.
'    On Error GoTo ErrHandle
    Dim NumberLength
    Dim Location
    Dim Checksum
    Dim Digit

    NumberLength = Len(sCard)

    'Add even digits in even length strings
    'or odd digits in odd length strings.
    For Location = 2 - (NumberLength Mod 2) To NumberLength Step 2
        Checksum = CInt(Mid(sCard, Location, 1)) + Checksum
    Next

    'Analyze odd digits in even length strings
    'or even digits in odd length strings.
    For Location = (NumberLength Mod 2) + 1 To NumberLength Step 2
        Digit = CInt(Mid(sCard, Location, 1)) * 2
        If Digit < 10 Then
            Checksum = Digit + Checksum
        Else
            Checksum = Digit - 9 + Checksum
        End If
    Next

    'Is the checksum divisible by ten?
    LuhnCheck = ((Checksum Mod 10) = 0)
End Function

Private CurrentScript	' sometimes we need to remember the current MT2OFXScript object

Class MT2OFXScript
	Public MinimumProgramVersion
' runtime
	Public DebugRecognition		' enables debug code in recognition
	Public AccountNum			' default if not specified in file
	Public BranchCode			' default if not specified in file
	Public AccountType			' defaults to CHECKING: can also be CREDITCARD etc

' specific to this format
	Public ScriptName
	Public FormatName
' 20091013 CS: add codepage to class
    Public CodePage
	Public ParseErrorMessage
	Public ParseErrorTitle
	Public BankCode
	Public QuickenBankID		' copied into INTU.BID if present
	Public CurrencyCode			' default if not specified in file
	Public OldestLast			' transactions are in reverse order

' callback function hooks
	Public ReadLineCallback		' GetRef() of function to call on all newly-read lines
	Public IsValidTxnLine		' GetRef() of function to call on newly-read line containing txn
	Public PreParseCallback		' GetRef() of function to call after line is split into fields
	Public TransactionCallback	' GetRef() of function to call in main script to post-process txn
	Public StatementCallback	' GetRef() of function to call in main script to post-process statement
	Public HeaderCallback		' GetRef() of function to call in main script to process a header line
	Public CustomDateCallback	' GetRef() of custom date parser
	Public CustomAmountCallback	' GetRef() of custom amount parser
	Public FinaliseCallback		' GetRef() of finalisaion callback
' 20090827 CS: add user field callbacks
    Public User1Callback
    Public User2Callback
    Public User3Callback
    Public User4Callback
    Public User5Callback
	Public MyReadLine			' GetRef() of function to read next line of input
	Public MyRewind				' GetRef() of function to reset input to beginning
	Public MyAtEOF				' GetRef() of function to detect EOF

' CSV parsing
	Public SkipHeaderLines		' number of lines to skip before the transaction data
	Public ColumnHeadersPresent	' are the column headers in the file?
	Public DecimalSeparator		' as used in amounts - "," or "."
' If CSVSeparator is empty, TxnLinePattern (RegExp pattern) is used to parse the line. The "fields" correspond
' to the text between the top-level parentheses.
	Public CSVSeparator			' usually "," or ";"
	Public TxnLinePattern
	Public MinFieldsExpected
	Public MaxFieldsExpected

' field parsing
	Public DateSequence			' must be DMY, MDY, or YMD
	Public DateSeparator		' multi-char string, or can be empty for dates in e.g. "yyyymmdd" format
	Public InvertSign			' make credits into debits etc
	Public NoAvailableBalance	' True if file does not contain "Available Balance" information
	Public MemoChunkLength		' if memo field consists of fixed length chunks
	Public TxnDatePattern		' RegExp pattern to find transaction date in the memo
	Public TxnDateSequence		' e.g. Array(3,2,1,4,5,0): order of the info in the pattern: Y,M,D,H,M,S
	Public PayeeLocation		' start of payee in memo
	Public PayeeLength			' length of payee in memo
	
	Private xFields				' field definition array
	Private xFieldDict
	Private xProperties			' property definition array

	Private xMonthNames			' month names in dates
	Private sLastLine			' whole line last returned by NextLine()

' MonthNames property
	Public Property Get MonthNames()
		MonthNames = xMonthNames
	End Property
	Public Property Let MonthNames(x)
		If IsArray(x) Then
			If (UBound(x) Mod 12) <> 11 Then
				Message True, True, ScriptName, "Length of MonthNames array must be a multiple of 12"
				Err.Raise 1, ScriptName, "Length of MonthNames array must be a multiple of 12"
				Exit Property
			End If
		End If
		xMonthNames = x
		InitialiseMonths xMonthNames
	End Property

' ScriptTitle property
	Public Property Get ScriptTitle()
		ScriptTitle = FormatName & " (" & ScriptName & ")"
	End Property
	
' Fields property
	Public Property Get Fields()
		Fields = xFields
	End Property
	Public Property Let Fields(x)
		xFields = x
		MaxFieldsExpected = UBound(x)+1
		MinFieldsExpected = UBound(x)+1
' fill field lookup dictionary
' NB: only the last occurrence is remembered!
      xFieldDict.RemoveAll
		Dim i
		For i=0 To UBound(xFields)
			xFieldDict(xFields(i)(0)) = i+1
		Next
	End Property

' Properties property
	Public Property Get Properties()
		Properties = xProperties
	End Property
	Public Property Let Properties(x)
		xProperties = x
		If IsArray(xProperties) Then
			LoadProperties ScriptName, xProperties
		End If
	End Property

' FieldDict property (read-only)
	Public Function FieldDict()
		Set FieldDict = xFieldDict
	End Function

' GetTransactionDate method
	Public Function GetTransactionDate(t)
		Dim sMemo: sMemo = t.Memo
		Dim vDateBits
		Dim dTxn
		Dim iYear, iMonth, iDay, iHour, iMin, iSec
		If t.TxnDate <> NODATE Then
			GetTransactionDate = t.TxnDate
			Exit Function
		End If
		dTxn = NODATE
		If Len(TxnDatePattern) = 0 Then
			Err.Raise 1, ScriptName, "GetTransactionDate called with empty pattern"
			GetTransactionDate = dTxn
			Exit Function
		End If
		vDateBits = ParseLineFixed(sMemo, TxnDatePattern)
		If TypeName(vDateBits) = "Variant()" Then
        If UBound(vDateBits) >= 0 Then
			If TxnDateSequence(0) > 0 Then iYear = CInt(vDateBits(TxnDateSequence(0)))
			If TxnDateSequence(1) > 0 Then iMonth = CInt(vDateBits(TxnDateSequence(1)))
			If TxnDateSequence(2) > 0 Then iDay = CInt(vDateBits(TxnDateSequence(2)))
			If TxnDateSequence(3) > 0 Then iHour = CInt(vDateBits(TxnDateSequence(3)))
			If TxnDateSequence(4) > 0 Then iMin = CInt(vDateBits(TxnDateSequence(4)))
			If TxnDateSequence(5) > 0 Then iSec = CInt(vDateBits(TxnDateSequence(5)))
			dTxn = DateSerial(iYear, iMonth, iDay) + TimeSerial(iHour, iMin, iSec)
         End If
      Else
'			MsgBox "GetTransactionDate: " & sMemo & " parsed to " & TypeName(vDateBits)
		End If
		GetTransactionDate = dTxn
	End Function

' ParseDate method
	Public Function ParseDate(sDate)
		If IsEmpty(CustomDateCallback) Then
			ParseDate = ParseDateEx(sDate, DateSequence, DateSeparator)
		Else 
			ParseDate = CustomDateCallback(sDate)
		End If
	End Function

' ParseAmount method
	Public Function ParseAmount(sAmt)
		If IsEmpty(CustomAmountCallback) Then
			ParseAmount = ParseNumber(sAmt, DecimalSeparator)
		Else
			ParseAmount = CustomAmountCallback(sAmt)
		End If
	End Function

' NextLine method
	Public Function NextLine()
		Dim sLine
		sLine = MyReadLine()
		sLine = ReadLineCallback(sLine)
		NextLine = sLine
		sLastLine = sLine
	End Function

' LastLine method
	Public Function LastLine()
		LastLine = sLastLine
	End Function
	
' NoMoreInput method
	Public Function NoMoreInput()
		NoMoreInput = MyAtEOF()
	End Function

' MatchColumnHeaders method
	Public Function MatchColumnHeaders(vFields)
		Dim i, sField, sTmp
		MatchColumnHeaders = False
		For i=LBound(vFields) To UBound(vFields)
			sField = Trim(vFields(i))
			If UBound(xFields(i-1)) > 1 Then
				If Left(xFields(i-1)(2), 1) = "=" Then
					sTmp = Replace(Mid(xFields(i-1)(2), 2), "%1", sField)
					If Not Eval(sTmp) Then
						If DebugRecognition Then
							MsgBox "Field " & CStr(i) & ": '" & sField & "' failed custom match '" & xFields(i-1)(2) & "'",,ScriptName
						End If
						Exit Function
					End If
				Else
					If Not StringMatches(sField, xFields(i-1)(2)) Then
						If DebugRecognition Then
							MsgBox "Field " & CStr(i) & ": '" & sField & "' does not match '" & xFields(i-1)(2) & "'",,ScriptName
						End If
						Exit Function
					End If
				End If
			Else
				If sField <> xFields(i-1)(1) Then
					If DebugRecognition Then
						MsgBox "Field " & CStr(i) & " " & sField & ", expecting " & xFields(i-1)(1),,ScriptName
					End If
					Exit Function
				End If
			End If
		Next
		MatchColumnHeaders = True
	End Function
	
' MatchTransactionPattern method
	Public Function MatchTransactionPattern(vFields)
		Dim i, sField, sPat, bTmp, sTmp
' pattern-match the first row
		For i=LBound(vFields) To UBound(vFields)
			sField = Trim(vFields(i))
			If UBound(xFields(i-1)) > 1 Then
				sPat = xFields(i-1)(2)
				If Left(sPat, 1) = "=" Then
' this is going to cause a problem as functions in a script module are not exposed to Global
					sTmp = Replace(Mid(sPat, 2), "%1", sField)
					bTmp = Eval(sTmp)
				Else
					bTmp = StringMatches(sField, sPat)
				End If
			Else
				Select Case xFields(i-1)(0)
				case fldSkip, fldMemo, fldPayee, fldCategory
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
				Case fldSign
					sPat = "\+|-|CR|DR|C|D"
					bTmp = StringMatches(sField, sPat)
            Case Else
                bTmp = True
				End Select
			End If
			If Not bTmp Then
				If DebugRecognition Then
					MsgBox "Field " & i & " (" & sField & ") failed to match '" & sPat & "'",,ScriptName
				End If
				Exit For
			End If
		Next
		MatchTransactionPattern = bTmp
	End Function

    Public Function DoUserFieldCallback(iField, stmt, txn, sField)
    Dim bTmp
        Select Case iField
        Case fldUser1
            bTmp = User1Callback(stmt, txn, sField)
        Case fldUser2
            bTmp = User2Callback(stmt, txn, sField)
        Case fldUser3
            bTmp = User3Callback(stmt, txn, sField)
        Case fldUser4
            bTmp = User4Callback(stmt, txn, sField)
        Case fldUser5
            bTmp = User5Callback(stmt, txn, sField)
        Case Else
            bTmp = False
        End Select
        DoUserFieldCallback = bTmp
    End Function

' Initialisation and termination
	Private Sub Class_Initialize
		Set xFieldDict = CreateObject("Scripting.Dictionary")
		Set ReadLineCallback = GetRef("DummyCallback1a")
		Set IsValidTxnLine = GetRef("IsValidTxnLineDefault")
		Set PreParseCallback = GetRef("DummyCallback1")
		Set TransactionCallback = GetRef("DummyCallback2")
		Set StatementCallback = GetRef("DummyCallback1")
		Set HeaderCallback = GetRef("DummyCallback1")
		Set FinaliseCallback = GetRef("DummyCallback0")
      Set User1Callback = GetRef("DummyCallbackUserField")
      Set User2Callback = GetRef("DummyCallbackUserField")
      Set User3Callback = GetRef("DummyCallbackUserField")
      Set User4Callback = GetRef("DummyCallbackUserField")
      Set User5Callback = GetRef("DummyCallbackUserField")
		Set MyReadLine = GetRef("StdReadLine")
		Set MyRewind = GetRef("StdRewind")
		Set MyAtEOF = GetRef("StdAtEOF")
' 20091013 CS: add codepage to class
        CodePage = CP_ACP
		AccountType = "CHECKING"
	End Sub

	Private Sub Class_Terminate
	End Sub
End Class

' default functions for use with callbacks
	Public Function DummyCallback0()
		DummyCallback0 = True
	End Function
	Public Function DummyCallback1(x)
		DummyCallback1 = True
	End Function
	Public Function DummyCallback1a(x)
		DummyCallback1a = x
	End Function
	Public Function DummyCallback2(x,y)
		DummyCallback2 = True
	End Function
    Public Function DummyCallbackUserField(s,t,v)
        DummyCallbackUserField = True
    End Function
	Private Function IsValidTxnLineDefault(sLine)
		IsValidTxnLineDefault = (Len(sLine) > 0) And Left(sLine,1) <> CurrentScript.CSVSeparator
	End Function
	Private Sub StdRewind()
		Rewind
	End Sub
	Private Function StdReadLine()
		StdReadLine = ReadLine()
	End Function
	Private Function StdAtEOF()
		StdAtEOF = AtEOF()
	End Function
	
Function DefaultRecogniseTextFile(p)
	Dim vFields
	Dim sLine
	Dim i
	Dim bTmp
	Dim sField
	Dim sPat
	Dim nFields
	DefaultRecogniseTextFile = False
	Set CurrentScript = p

	If Len(p.MinimumProgramVersion) > 0 Then
		If Not VersionAtLeast(p.MinimumProgramVersion) Then
			MsgBox "This MT2OFX script requires at least " & p.MinimumProgramVersion & " of the program and you have version " & Version & ".", _
				vbOKOnly+vbInformation, p.ScriptName
			Abort
			Exit Function
		End If
	End If

' 20091013 CS: codepage now in class
    If Session.InputFile.CodePage <> CP_ACP Then
        Session.InputFile.CodePage = p.CodePage
    End If
    
	For i=1 To p.SkipHeaderLines
		If p.MyAtEOF() Then
			Exit Function
		End If
		sLine = p.NextLine()
		p.HeaderCallback(sLine)
	Next
	If p.NoMoreInput() Then
		Exit Function
	End If
	sLine = p.NextLine()
	If p.ColumnHeadersPresent Then
		p.HeaderCallback(sLine)
	End If
' 20090827 CS: auto detect separator
	If Len(p.TxnLinePattern) > 0 Then
		vFields = ParseLineFixed(sLine, p.TxnLinePattern)
	Else
        If Len(p.CSVSeparator) = 0 Then
            p.CSVSeparator = FindCSVSeparator(sLine)
            If Len(p.CSVSeparator) = 0 Then
                Message True, True, "Unable to auto-detect CSV separator in '" & sLine & "'", p.ScriptName
                Abort
                Exit Function
            End If
        End If
		vFields = ParseLineDelimited(sLine, p.CSVSeparator)
	End If
	If TypeName(vFields) <> "Variant()" Then
		If p.DebugRecognition Then
			Message True, True, "Line not returned as array", p.ScriptName
		End If
		Err.Raise 1, p.ScriptName, "Line not returned as array"
		Exit Function
	End If
	nFields = UBound(vFields) - LBound(vFields) + 1
	If nFields < p.MinFieldsExpected Or nFields > p.MaxFieldsExpected Then
		If p.DebugRecognition Then
			Message True, True, "Wrong number of fields - got " & CStr(nFields) _
				& ", expected " & p.MinFieldsExpected & "-" & p.MaxFieldsExpected & VbCrLf & sLine, _
				p.ScriptName
		End If
		Exit function
	End If
	p.PreParseCallback(vFields)
	If p.ColumnHeadersPresent Then	' match against the headers
		DefaultRecogniseTextFile = p.MatchColumnHeaders(vFields)
	Else
		DefaultRecogniseTextFile = p.MatchTransactionPattern(vFields)
	End If
End Function

Function CanProcessLine()
End Function

'===================================================================
' Class ProcessState
' maintains state for DefaultLoadTextFile between transactions For
' interleaved statements such as PayPal and Yodlee
'
Class ProcessState
	Dim Stmt	' statement under construction
	Dim da		' date accumulator for this statement
' Initialisation and termination
	Private Sub Class_Initialize
		Set da = New DateAccumulator
	End Sub
	Private Sub Class_Terminate
	End Sub	
End Class

Private xStates: Set xStates=CreateObject("Scripting.Dictionary")

Public Function FindState(sKey)
	If xStates.Exists(sKey) Then
		Set FindState = xStates(sAcct)
	Else
		Set FindState = Nothing
	End If
End Function

Public Function GetState(sKey)
	Dim x
	Set x = FindState(sKey)
	If x Is Nothing Then
		Set x = New ProcessState
		Set xStates(sKey) = x
		Set x.Stmt = NewStatement()
'		x.Stmt.Acct = sAcct
	End If
	Set GetState = x
End Function
'========================================================================

Function DefaultLoadTextFile(p)
	Dim sLine       ' holds a line
	Dim vFields     ' array of fields in the line
	Dim pFields		 ' current field definitions
   Dim iField      ' current field code
	Dim sAcct       ' last account number
	Dim Stmt        ' holds the current statement
	Dim sTmp		    ' temporary string
	Dim i
	Dim dBal		   ' temp balance date
	Dim sField		' field value being processed
	Dim FirstTxn
	Dim sSign		' + or - (if present)
	Dim daTxnDate: Set daTxnDate = New DateAccumulator
	Dim nFields
	Dim st			' current ProcessState
   Dim tlType     ' transaction line type
   Dim bNewStatement   ' need to set up a new statement
	
    DefaultLoadTextFile = False
    Set CurrentScript = p

' 20091013 CS: codepage now in class
    If Session.InputFile.CodePage <> CP_ACP Then
        Session.InputFile.CodePage = p.CodePage
    End If
    
    For i=1 To p.SkipHeaderLines
        If p.NoMoreInput() Then
            Exit Function
        End If
        sLine = p.NextLine()
        p.HeaderCallback(sLine)
    Next
    If p.NoMoreInput() Then
        Exit Function
    End If
    If p.ColumnHeadersPresent Then
        sLine = p.NextLine()
        p.HeaderCallback(sLine)
    End If

    Do While Not p.NoMoreInput()
        sLine = p.NextLine()
        tlType = CInt(p.IsValidTxnLine(sLine)) ' translate True/False for old scripts
        If tlType <> txnlineSKIP Then

' 20090827 CS: auto detect separator
            If Len(p.TxnLinePattern) > 0 Then
                vFields = ParseLineFixed(sLine, p.TxnLinePattern)
            Else
                If Len(p.CSVSeparator) = 0 Then
                    p.CSVSeparator = FindCSVSeparator(sLine)
                    If Len(p.CSVSeparator) = 0 Then
                        Message True, True, "Unable to auto-detect CSV separator in '" & sLine & "'", p.ScriptName
                        Abort
                        Exit Function
                    End If
                End If
                vFields = ParseLineDelimited(sLine, p.CSVSeparator)
            End If

            If TypeName(vFields) <> "Variant()" Then
                MsgBox "Parse Error: '" & sLine & "'", vbOkOnly+vbCritical, "Parse Error"
                Abort
                Exit Function
            End If
            nFields = UBound(vFields) - LBound(vFields) + 1
            If nFields < p.MinFieldsExpected Or nFields > p.MaxFieldsExpected Then
                Message True, True, "Wrong number of fields - got " & CStr(nFields) _
                    & ", expected " & p.MinFieldsExpected & "-" & p.MaxFieldsExpected & VbCrLf & sLine, p.ScriptName
        msgbox p.TxnLinePattern
                Abort
                Exit Function
            End If
         
            If p.PreParseCallback(vFields) Then
         
' should we start a new statement here?
            bNewStatement = False
            If IsEmpty(Stmt) Then
                bNewStatement = True
'msgbox "New statement due to no current statement"
                sAcct = p.AccountNum
            ElseIf p.FieldDict.Exists(fldAccountNum) Then
                If Stmt.Acct <> vFields(p.FieldDict.Item(fldAccountNum)) Then
                    If Stmt.Acct = "" Then
'msgbox "Setting account number from file to " & vFields(p.FieldDict.Item(fldAccountNum))
                        Stmt.Acct = vFields(p.FieldDict.Item(fldAccountNum))
                    Else
                        sAcct = vFields(p.FieldDict.Item(fldAccountNum))
'msgbox "New statement due to change in account number present in file"
                        bNewStatement = True
                    End If
                End If
            ElseIf Stmt.Acct <> p.AccountNum Then
                If Stmt.Acct = "" And p.AccountNum <> "" Then
'msgbox "Setting account num from parameters to " & p.AccountNum
                    Stmt.Acct = p.AccountNum
                ElseIf p.AccountNum <> "" then
                    sAcct = p.AccountNum
'msgbox "New statement due to change in account number in parameters from " & Stmt.Acct & " to " & p.AccountNum
                    bNewStatement = True
                End If
            ElseIf tlType = txnlineNORMALNEWSTATEMENT Or tlType = txnlineNOTRANSACTIONNEWSTATEMENT Then
                bNewStatement = True   ' script forcing a new statement
            End If

' if needed set up a new statement
            If bNewStatement Then
                PostprocessStatement p, Stmt
'msgbox "New statement for " & sAcct
                Set Stmt = NewStatement()
                Stmt.OpeningBalance.Ccy = p.CurrencyCode
                If Not p.NoAvailableBalance Then Stmt.AvailableBalance.Ccy = p.CurrencyCode
                Stmt.ClosingBalance.Ccy = p.CurrencyCode
                Stmt.BankName = p.BankCode
                Stmt.Acct = sAcct
                Stmt.BranchName = p.BranchCode
                Stmt.AcctType = p.AccountType
                FirstTxn = True               
            End If
            
            If tlType = txnlineNORMAL Or tlType = txnlineNORMALNEWSTATEMENT Then
                NewTransaction
            ElseIf tlType = txnlineCONTINUE Then
                If FirstTxn Then
                    Message True, True, p.ScriptName, "Continuation line with no active transaction"
                    Abort
                End If
            End If
            
            sSign = "+"
            LastMemo = ""
            For i=LBound(vFields) To UBound(vFields)
                sField = Trim(vFields(i))
                pFields = p.Fields
                iField = pFields(i-1)(0)
                Select Case iField
                case fldSkip, fldEmpty
                    ' do nothing
                case fldAccountNum
                    Stmt.Acct = sField
                case fldBranch
                    Stmt.BranchName = sField
                case fldCurrency
                    Stmt.OpeningBalance.Ccy = sField
                    Stmt.ClosingBalance.Ccy = sField
                    If Not p.NoAvailableBalance Then Stmt.AvailableBalance.Ccy = sField
                Case fldClosingBal
                    If p.OldestLast Then
                        If FirstTxn Then
                            Stmt.ClosingBalance.Amt = p.ParseAmount(sField)
                        End If
                    Else
                        Stmt.ClosingBalance.Amt = p.ParseAmount(sField)
                    End If
                case fldAvailBal
                    Stmt.AvailableBalance.Amt = p.ParseAmount(sField)
                case fldBookDate
                    Txn.BookDate = p.ParseDate(sField)
                    daTxnDate.Process Txn.BookDate
                case fldValueDate
                    Txn.ValueDate = p.ParseDate(sField)
                    daTxnDate.Process Txn.ValueDate
                Case fldTransactionDate
                    Txn.TxnDate = p.ParseDate(sField)
                    Txn.TxnDateValid = (Txn.TxnDate <> NODATE)
                    daTxnDate.Process Txn.TxnDate
                Case fldTransactionTime
                    If Txn.TxnDate <> NODATE then
                        If Len(sField)=5 Then
                            Txn.TxnDate = Txn.TxnDate + TimeSerial(CInt(Left(sField,2)), _
                                CInt(Mid(sField,4,2)),0)
                        ElseIf Len(sField)=8 Then
                            Txn.TxnDate = Txn.TxnDate + TimeSerial(CInt(Left(sField,2)), _
                                CInt(Mid(sField,4,2)),CInt(Mid(sField,7,2)))
                        End If
                    End If
                case fldAmtCredit
                    Txn.Amt = Txn.Amt + Abs(p.ParseAmount(sField))
                case fldAmtDebit
                    Txn.Amt = Txn.Amt - Abs(p.ParseAmount(sField))
                Case fldAmount
                    Txn.Amt = p.ParseAmount(sField)
                Case fldChequeNum
                    Txn.CheckNum = sField
                case fldMemo
                    ConcatMemo sField
                Case fldBalanceDate
                    dBal = p.ParseDate(sField)
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
                Case fldSign
                    sSign = sField
                Case fldCategory
                    Txn.Category = sField
                Case fldPayeeCity
                    Txn.Payee.City = sField
                Case fldPayeeState
                    Txn.Payee.State = sField
                Case fldPayeeZip
                    Txn.Payee.PostalCode = sField
                Case fldPayeeCountry
                    Txn.Payee.Country = sField
                Case fldPayeePhone
                    Txn.Payee.Phone = sField
                Case fldPayeeAddress1
                    Txn.Payee.Addr1 = sField
                Case fldPayeeAddress2
                    Txn.Payee.Addr2 = sField				
                Case fldPayeeAddress3
                    Txn.Payee.Addr3 = sField				
                Case fldPayeeAddress4
                    Txn.Payee.Addr4 = sField				
                Case fldPayeeAddress5
                    Txn.Payee.Addr5 = sField				
                Case fldPayeeAddress6
                    Txn.Payee.Addr6 = sField
' 20080827 CS: add user fields
                Case fldUser1, fldUser2, fldUser3, fldUser4, fldUser5
                    p.DoUserFieldCallback iField, Stmt, Txn, sField
' 20091005 CS: add payee account fields
                Case fldPayeeAcctNum
                    Txn.Payee.Acct = sField
                Case fldPayeeAcctBank
                    Txn.Payee.BankName = sField
                Case fldPayeeAcctBranch
                    Txn.Payee.BranchName = sField
                Case fldPayeeAcctType
                    Txn.Payee.AcctType = sField
                Case fldPayeeAcctKey
                    Txn.Payee.AcctKey = sField
                End Select
            Next
        
' correct the sign of the amount
            If tlType <> txnlineNOTRANSACTION And tlType <> txnlineNOTRANSACTIONNEWSTATEMENT Then
                If sSign = "-" Or sSign = "DR" Or sSign = "D" Then
                    Txn.Amt = -Txn.Amt
                End If
                If p.InvertSign Then
                    Txn.Amt = -Txn.Amt
                End If
            End If

' transaction type
            If tlType <> txnlineNOTRANSACTION And tlType <> txnlineNOTRANSACTIONNEWSTATEMENT Then
                If Txn.Amt < 0 Then
                    Txn.TxnType = "PAYMENT"
                Else
                    Txn.TxnType = "DEP"
                End If
            End If
			
            Dim sMemo
' find the payee, transaction type and txn date if we can
            If tlType <> txnlineNOTRANSACTION And tlType <> txnlineNOTRANSACTIONNEWSTATEMENT Then
                sMemo = Txn.Memo
                If p.PayeeLocation > 0 And Len(Txn.Payee) = 0 Then
                    Txn.Payee = Trim(Mid(sMemo, p.PayeeLocation, p.PayeeLength))
                End If
                If Len(p.TxnDatePattern) > 0 Then
                    Txn.TxnDate = p.GetTransactionDate(Txn)
                    If Txn.TxnDate <> NODATE Then
                        Txn.TxnDateValid = True
                        daTxnDate.Process Txn.TxnDate
                    End If
                End If
            End If
        
' tidy up the memo
            If tlType <> txnlineNOTRANSACTION And tlType <> txnlineNOTRANSACTIONNEWSTATEMENT Then
                If p.MemoChunkLength > 0 Then
                    sMemo = Txn.Memo
                    Txn.Memo = ""
                    For i=1 To Len(sMemo) Step p.MemoChunkLength
                        ConcatMemo Trim(Mid(sMemo, i, p.MemoChunkLength))
                    Next
                End If
            End If

' postprocess the transaction
            If tlType <> txnlineNOTRANSACTION And tlType <> txnlineNOTRANSACTIONNEWSTATEMENT Then
                Call p.TransactionCallback(Txn, vFields)
            End If
			
' keep tabs on the statement/balance Date
            Stmt.ClosingBalance.BalDate = daTxnDate.MaxDate
            Stmt.OpeningBalance.BalDate = daTxnDate.MinDate
			
            If tlType <> txnlineNOTRANSACTION And tlType <> txnlineNOTRANSACTIONNEWSTATEMENT Then
                FirstTxn = False
            End If
        End If
		End If
    Loop

' postprocess the final statement
    PostprocessStatement p, Stmt

' postprocess the whole business
    Bcfg.IntuitBankID = p.QuickenBankID
    DefaultLoadTextFile = p.FinaliseCallback()
End Function

Private Function PostprocessStatement(p, s)
    Dim sTmp
    If Not IsEmpty(s) Then
        Call p.StatementCallback(s)
        sTmp = MapAccount(s.Acct)
        If Len(sTmp) > 0 Then
            s.Acct = sTmp
        End If
    End If
End Function

Public Function PropertyExists(sProp)
	PropertyExists = False
	If IsEmpty(CurrentScript.Properties) Then
		Exit Function
	End If
	If Not IsArray(CurrentScript.Properties) Then
		Exit Function
	End If
	Dim p
	For Each p In CurrentScript.Properties
		If IsArray(p) Then
			If p(0) = sProp Then
				PropertyExists = True
				Exit Function
			End If
		End If
	Next
End Function

Public Function Calculate_Mod97(sAcct)
	Dim sTmp, iCheck, iTmp
	Dim sChunk, sRem
	sRem = ""
	sTmp = sAcct
	Do While Len(sTmp) > 0
		sChunk = sRem & Left(sTmp, (9-Len(sRem)))
		sTmp = Mid(sTmp, 9-Len(sRem)+1)
		iTmp = CLng(sChunk) Mod 97
		sRem = Right("0" & CStr(iTmp), 2)
	Loop
	Calculate_Mod97 = iTmp
End Function

Public Function Validate_Mod97(sAcct)
	Dim sTmp, iTmp, iCheck
	sTmp = ExtractDigits(sAcct)
	iCheck = CInt(Right(sTmp, 2))
	iTmp = Calculate_Mod97(Left(sTmp, Len(sTmp)-2))
	Validate_Mod97 = (iTmp = iCheck)
End Function

Public Function Validate_Belgium(sAcct)
	Dim sTmp
	sTmp = ExtractDigits(sAcct)
	If Len(sTmp) <> 12 Then
		Exit Function
	End If
	Validate_Belgium = Validate_Mod97(sTmp)
End Function

Public Function Validate_Netherlands(sAcct)
    Dim sTmp, i, iSum
    Validate_Netherlands = False
    sTmp = ExtractDigits(sAcct)
	If Len(sTmp) <> 9 And Len(sTmp) <> 10 Then
		Exit Function
	End If
	iSum = 0
	If Len(sTmp) = 10 Then
		If Left(sTmp,1) <> "0" Then
			Exit Function
		End If
		sTmp = Mid(sTmp, 2)
	End If
	For i=1 To 9
		iSum = iSum + (CInt(Mid(sTmp, i, 1)) * (i+1))
	Next
	Validate_Netherlands = ((iSum Mod 11) = 0)
End Function

Public Function Validate_IBAN(sAcct)
	Dim sTmp, i, c, ic, sTmp2
	Dim iA, iZ
' canonicalise input - upper case, no spaces
	sTmp = UCase(Replace(sAcct, " ", ""))
' step 1: move first four chars to the end
	sTmp = Mid(sTmp, 5) & Left(sTmp, 4)
' step 2: replace letters with equivalent digits
	sTmp2 = ""
	iA = Asc("A"): iZ = Asc("Z")
	For i=1 To Len(sTmp)
		c = Mid(sTmp, i, 1)
		ic = Asc(c)
		If ic>=iA And ic<=aZ Then
			sTmp2 = sTmp2 & CStr(ic-iA+10)
		Else
			sTmp2 = sTmp2 & c
		End If
	Next
' step 3: calculate mod97 check digits
	Validate_IBAN = (Calculate_Mod97(sTmp2) = 1)
End Function

' HashString returns an MD5 hash (base-64 encoded) of the input string. This function can be used
' to generate a FITID value based on any parts of the transaction; the standard algorithm
' covers Memo, Payee, Amount, BookDate.
Public Function HashString(sInput)
	Dim sTmp
	sTmp = ""
    If HashStart() Then
        If HashAddString(sInput) Then
        	sTmp = HashEnd()
        End If
    End If
    HashString = sTmp
End Function

Class DateAccumulator
	Private xMinDate
	Private xMaxDate
	Public Property Get MaxDate()
		MaxDate = xMaxDate
	End Property
	Public Property Get MinDate()
		MinDate = xMinDate
	End Property
	Public Sub Process(d)
		If d = NODATE Then
			Exit Sub
		End If
		If xMinDate = NODATE Or xMinDate > d Then
			xMinDate = d
		End If
		If xMaxDate = NODATE Or xMaxDate < d Then
			xMaxDate = d
		End If
	End Sub
	Public Sub Reset
		xMinDate = NODATE
		xMaxDate = NODATE
	End Sub
' Initialisation and termination
	Private Sub Class_Initialize
		Call Reset
	End Sub	
End Class
'GetConfigStringEx(sSection As Variant, sKey As Variant, sDefault As Variant, sFile As Variant) As String
Public Function MapAccount(sAcct)
	MapAccount = GetConfigStringEx("Accounts", sAcct, "", Cfg.AppDataPath & "\accounts.ini")
End Function
Public Function MapCountryToISO3166(sCountry)
	MapCountryToISO3166 = GetConfigStringEx("ToISO3166", sCountry, "", Cfg.AppPath & "\countries.ini")
End Function
