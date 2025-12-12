Attribute VB_Name = "OFXWrite"
Option Explicit

' $Header: /MT2OFX/OFXWrite.bas 20    24/11/09 22:05 Colin $

Private OutFile As OutputFile
Private Indent As Integer
Private sIndent As String
Private SpacesPerIndent As Integer
Private OFXHeader As Integer
Private OFXVersion As Integer
Private Const OFXDefaultVersion As Integer = 102
Private sMyAcctType As String   ' contains OFX ACCTTYPE or "CREDITCARD" or "STOCKPRICES"
Private Const csCreditCard As String = "CREDITCARD"
Private Const csStockPrices As String = "STOCKPRICES"
Private iLocalTZbias As Long

Public Function FormatOFXDateTime(dDate As Date) As String
    FormatOFXDateTime = Format(dDate, "yyyymmddHhNnSs")
    If Right$(FormatOFXDateTime, 6) = "000000" Then
        FormatOFXDateTime = Left$(FormatOFXDateTime, 8) & "120000"
    End If
End Function

Public Function OFXEscape(sLine As String) As String
    Dim sTmp As String
    Dim lChar As Long
    Dim cTmp As String
    Dim i As Long
' CS 20041107: OFX versions from 1.5 on are moving more and more towards XML.
' 1.5.1 says leading and trailing spaces are not preserved - CDATA must be used.
' Embedded spaces are OK. OFX 2.0+ is full XML.
    If OFXVersion >= 151 Then
        If Left$(sLine, 1) = " " Or Right$(sLine, 1) = " " Then
            OFXEscape = "<![CDATA[" & sLine & "]]>"
            Exit Function
        End If
    End If
' CS 20041107: Now using hand-coded XML encoding below so non-ASCII chars are
' fully supported.
    sTmp = ""
    For i = 1 To Len(sLine)
        cTmp = Mid$(sLine, i, 1)
        lChar = AscW(cTmp) And &HFFFF&
        If lChar < 32 Or lChar > 126 Then
            If OFXVersion >= 200 Then
                If OutFile.CharInCodepage(cTmp) Then
                    sTmp = sTmp & cTmp
                Else
                    sTmp = sTmp & "&#x" & Hex(lChar) & ";"
                End If
            Else
                If lChar < 32 Then  ' control chars - yuck!
                    sTmp = sTmp & "?"
                Else    ' hope the code page will sort it out...
                    sTmp = sTmp & cTmp
                End If
            End If
        Else
            If cTmp = "&" Then
                sTmp = sTmp & "&amp;"
            ElseIf cTmp = "<" Then
                sTmp = sTmp & "&lt;"
            ElseIf cTmp = ">" Then
                sTmp = sTmp & "&gt;"
            Else
                sTmp = sTmp & cTmp
            End If
        End If
    Next
    OFXEscape = sTmp
End Function
Public Sub OFXPrintLine(sLine As String)
    OutFile.PrintLine sLine
End Sub
Public Sub OFXOpenTag(sTag As String)
    OFXPrintLine sIndent & "<" & sTag & ">"
    Indent = Indent + 1
    sIndent = Space$(Indent * SpacesPerIndent)
End Sub
Public Sub OFXTag(sTag As String, sText As String)
    If OFXVersion >= 200 Then
        OFXPrintLine sIndent & "<" & sTag & ">" & OFXEscape(sText) & "</" & sTag & ">"
    Else
        OFXPrintLine sIndent & "<" & sTag & ">" & OFXEscape(sText)
    End If
End Sub
Public Sub OFXCloseTag(sTag As String)
    Indent = Indent - 1
    sIndent = Space$(Indent * SpacesPerIndent)
    OFXPrintLine sIndent & "</" & sTag & ">"
End Sub

' 20050220 CS: Add language code to session
Public Sub OFXFileHeader(oFile As OutputFile, dNow As Date, sIntuitID As String, _
    sLanguage As String)
    Dim sTmp As String
    If OFXHeader = 0 Then
        OFXSetVersion OFXDefaultVersion
    End If
    Set OutFile = oFile
    Indent = 0
    sIndent = ""
    SpacesPerIndent = 2
    iLocalTZbias = GetCurrentTimeBias() ' in minutes
    If OFXVersion >= 200 Then
        sTmp = XMLEncodingForCodepage(OutFile.CodePage)
        If Len(sTmp) = 0 Then
            OFXPrintLine "<?xml version=""1.0"" standalone=""no""?>"
        Else
            OFXPrintLine "<?xml version=""1.0"" encoding=""" & sTmp & """ standalone=""no""?>"
        End If
        OFXPrintLine "<?OFX OFXHEADER=""" & CStr(OFXHeader) & """" _
            & " VERSION=""" & CStr(OFXVersion) & """ SECURITY=""NONE"" OLDFILEUID=""NONE""" _
            & " NEWFILEUID=""NONE""?>"
    Else
        OFXPrintLine "OFXHEADER:" & CStr(OFXHeader)
        OFXPrintLine "DATA:OFXSGML"
        OFXPrintLine "VERSION:" & CStr(OFXVersion)
        OFXPrintLine "SECURITY:NONE"
        If OutFile.CodePage = CP_UTF8 Then
            OFXPrintLine "ENCODING:UTF-8"
            OFXPrintLine "CHARSET:1252"
        Else
            OFXPrintLine "ENCODING:USASCII"
            OFXPrintLine "CHARSET:" & CStr(OutFile.CodePage)
        End If
        OFXPrintLine "COMPRESSION:NONE"
        OFXPrintLine "OLDFILEUID:NONE"
        OFXPrintLine "NEWFILEUID:NONE"
    End If
    OFXPrintLine ""
    OFXOpenTag "OFX"
    OFXOpenTag "SIGNONMSGSRSV1"
    OFXOpenTag "SONRS"
    OFXOpenTag "STATUS"
    OFXTag "CODE", "0"
    OFXTag "SEVERITY", "INFO"
    OFXCloseTag "STATUS"
    OFXTag "DTSERVER", Format(dNow, "yyyymmddhhmmss")
    OFXTag "LANGUAGE", sLanguage
    If Len(sIntuitID) > 0 Then
        OFXTag "INTU.BID", sIntuitID
    End If
    OFXCloseTag "SONRS"
    OFXCloseTag "SIGNONMSGSRSV1"
End Sub

Public Sub OFXStatementSetHeader(oFile As OutputFile, sAcctType As String)
    Set OutFile = oFile
    sMyAcctType = sAcctType
    Select Case sMyAcctType
    Case "CREDITCARD"
        OFXOpenTag "CREDITCARDMSGSRSV1"
    Case csStockPrices
        OFXOpenTag "SECLISTMSGSRSV1"
    Case Else
        OFXOpenTag "BANKMSGSRSV1"
    End Select
End Sub
Public Sub OFXStatementSetTrailer(oFile As OutputFile)
    Set OutFile = oFile
    Select Case sMyAcctType
    Case csCreditCard
        OFXCloseTag "CREDITCARDMSGSRSV1"
    Case csStockPrices
        OFXCloseTag "SECLISTMSGSRSV1"
    Case Else
        OFXCloseTag "BANKMSGSRSV1"
    End Select
End Sub
' 20041208 CS Added support for Branch ID
Public Sub OFXStatementHeader(oFile As OutputFile, _
    sDefCur As String, sBankName As String, sBranch As String, sAcctNum As String, _
    sAcctType As String, dFrom As Date, dTo As Date)
    Set OutFile = oFile
    Select Case sMyAcctType
    Case csCreditCard
        OFXOpenTag "CCSTMTTRNRS"
    Case csStockPrices
        Exit Sub
    Case Else
        OFXOpenTag "STMTTRNRS"
    End Select
        OFXTag "TRNUID", "1"
        OFXOpenTag "STATUS"
            OFXTag "CODE", "0"
            OFXTag "SEVERITY", "INFO"
        OFXCloseTag "STATUS"
        Select Case sMyAcctType
        Case csCreditCard
            OFXOpenTag "CCSTMTRS"
        Case Else
            OFXOpenTag "STMTRS"
        End Select
            OFXTag "CURDEF", sDefCur
            Select Case sMyAcctType
            Case csCreditCard
                OFXOpenTag "CCACCTFROM"
                    OFXTag "ACCTID", sAcctNum
                OFXCloseTag "CCACCTFROM"
            Case Else
                OFXOpenTag "BANKACCTFROM"
                    OFXTag "BANKID", sBankName
                    If Len(sBranch) > 0 Then
                        OFXTag "BRANCHID", sBranch
                    End If
                    OFXTag "ACCTID", sAcctNum
                    OFXTag "ACCTTYPE", sAcctType
                OFXCloseTag "BANKACCTFROM"
            End Select
            OFXOpenTag "BANKTRANLIST"
                OFXTag "DTSTART", FormatOFXDateTime(dFrom)
                OFXTag "DTEND", FormatOFXDateTime(dTo)
End Sub

' 20050918 CS: Add SIC to signature
' 20051013 CS: Payee is now a clever class
Public Sub OFXTransaction(oFile As OutputFile, sTrnType As String, _
    dValDate As Date, dTrnDate As Date, dBookDate As Date, _
    bUseTrnDate As Boolean, aAmount As Double, sFITID As String, _
    sCheckNum As String, xPayee As Payee, sMemo As String, iSIC As Long)
    Set OutFile = oFile
    Dim sAmt As String
    sAmt = FormatAmount(aAmount, Cfg.OFXDecimal)
    OFXOpenTag "STMTTRN"
    OFXTag "TRNTYPE", sTrnType
    OFXTag "DTPOSTED", FormatOFXDateTime(dBookDate)
    If bUseTrnDate Then
        OFXTag "DTUSER", FormatOFXDateTime(dTrnDate)
    End If
    If dValDate <> NODATE Then
        OFXTag "DTAVAIL", FormatOFXDateTime(dValDate)
    End If
    OFXTag "TRNAMT", sAmt
    OFXTag "FITID", sFITID
' CS 20040826: Allow CHECKNUM to be optional. See also DoProcessTxn()
    If Len(sCheckNum) > 0 Then
        OFXTag "CHECKNUM", OFXString(sCheckNum, 12)
    End If
' 20050918 CS: Add SIC
    If iSIC > 0 Then
        OFXTag "SIC", CStr(iSIC)
    End If
' 20051013 CS: Payee name is now more complex
    OFXPayee xPayee
' 20091005 CS: output payee bank account data
    If xPayee.AcctType = "CREDITCARD" Then
        If Len(xPayee.Acct) > 0 Then
            OFXOpenTag "CCACCTTO"
            OFXTag "ACCTID", OFXString(xPayee.Acct, 22)
            If Len(xPayee.AcctKey) > 0 Then OFXTag "ACCTKEY", OFXString(xPayee.AcctKey, 22)
            OFXCloseTag "CCACCTTO"
        End If
    Else
        If Len(xPayee.Acct) > 0 And Len(xPayee.BankName) > 0 Then
            OFXOpenTag "BANKACCTTO"
            OFXTag "BANKID", OFXString(xPayee.BankName, 9)  ' mandatory!
            If Len(xPayee.BranchName) > 0 Then OFXTag "BRANCHID", OFXString(xPayee.BranchName, 22)
            OFXTag "ACCTID", OFXString(xPayee.Acct, 22)
            OFXTag "ACCTTYPE", IIf(Len(xPayee.AcctType) = 0, "CHECKING", xPayee.AcctType)
            If Len(xPayee.AcctKey) > 0 Then OFXTag "ACCTKEY", OFXString(xPayee.AcctKey, 22)
            OFXCloseTag "BANKACCTTO"
        End If
    End If
    If Len(sMemo) > 0 Then
        OFXTag "MEMO", sMemo
    End If
    OFXCloseTag "STMTTRN"
End Sub
Private Sub OFXPayee(xPayee As Payee)
    Dim sTmp As String
    If xPayee.IsSimple Then
 ' CS 20050109: Make NAME optional
        If Len(xPayee.Name) > 0 Then
            OFXTag "NAME", OFXString(xPayee.Name, 32)
        End If
    Else
        OFXOpenTag "PAYEE"
        If Len(xPayee.Name) > 0 Then
            OFXTag "NAME", OFXString(xPayee.Name, 32)
        End If
        If OFXVersion >= 210 Then
            If Len(xPayee.Name) > 32 Then
                OFXTag "EXTDNAME", OFXString(xPayee.Name, 100)
            End If
        End If
        If Len(xPayee.Addr1) > 0 Then
            OFXTag "ADDR1", OFXString(xPayee.Addr1, 32)
            If Len(xPayee.Addr2) > 0 Then
                OFXTag "ADDR2", OFXString(xPayee.Addr2, 32)
                If Len(xPayee.Addr3) > 0 Then
                    OFXTag "ADDR3", OFXString(xPayee.Addr3, 32)
                End If
            End If
        End If
        OFXTag "CITY", OFXString(xPayee.City, 32)
        OFXTag "STATE", OFXString(xPayee.State, 5)
        OFXTag "POSTALCODE", OFXString(xPayee.PostalCode, 11)
        If Len(xPayee.Country) > 0 Then
            OFXTag "COUNTRY", OFXString(xPayee.Country, 3)
        End If
        OFXTag "PHONE", OFXString(xPayee.Phone, 32)
        OFXCloseTag "PAYEE"
    End If
End Sub

Private Sub OFXSecID(sID As String, sType As String)
    OFXOpenTag "SECID"
    OFXTag "UNIQUEID", OFXString(sID, 32)
    OFXTag "UNIQUEIDTYPE", OFXString(sType, 10)
    OFXCloseTag "SECID"
End Sub
Public Sub OFXStockInfo(s As SecurityInfo, sBaseCurrency As String)
    OFXOpenTag "STOCKINFO"
    OFXSecID s.UniqueID, s.UniqueIDType
    OFXTag "SECNAME", OFXString(s.Name, 120)
    If Len(s.Ticker) > 0 Then OFXTag "TICKER", OFXString(s.Ticker, 32)
    If Len(s.FIID) > 0 Then OFXTag "FIID", OFXString(s.FIID, 32)
    If Len(s.Rating) > 0 Then OFXTag "RATING", OFXString(s.Rating, 10)
    If s.UnitPrice <> 0 Then OFXTag "UNITPRICE", FormatAmount(s.UnitPrice, Cfg.OFXDecimal)
    If s.UnitPriceDate <> NODATE Then OFXTag "DTASOF", FormatOFXDateTime(s.UnitPriceDate)
    If Len(s.Ccy) > 0 And s.Ccy <> sBaseCurrency Then
        OFXOpenTag "CURRENCY"
        OFXTag "CURRATE", FormatAmount(s.ExchangeRate, Cfg.OFXDecimal)
        OFXTag "CURSYM", OFXString(s.Ccy, 3)
        OFXCloseTag "CURRENCY"
    End If
    If Len(s.Memo) > 0 Then OFXTag "MEMO", OFXString(s.Memo, 120)
    OFXCloseTag "STOCKINFO"
End Sub
Public Sub OFXSecurityList(sl As SecurityList, sBaseCurrency As String)
    Dim X As SecurityInfo
    OFXOpenTag "SECLIST"
    For Each X In sl
        Select Case X.SecurityType
        Case sctyStock
            OFXStockInfo X, sBaseCurrency
        Case Else
            LogMessage False, True, "Unimplemented security type: " & X.SecurityType
        End Select
    Next
    OFXCloseTag "SECLIST"
End Sub
Private Function OFXString(s As String, iMaxLen As Long) As String
    Dim sTmp As String
    sTmp = Left$(Trim$(s), iMaxLen)
    If Len(sTmp) = 0 Then sTmp = "-"
    OFXString = sTmp
End Function
Public Sub OFXStatementTrailer(oFile As OutputFile, aLBal As Double, _
    dLBal As Date, bABalPresent As Boolean, aABal As Double, dABal As Date)
    Set OutFile = oFile
    If sMyAcctType = csStockPrices Then
        Exit Sub
    End If
    OFXCloseTag "BANKTRANLIST"
    OFXOpenTag "LEDGERBAL"
    OFXTag "BALAMT", FormatAmount(aLBal, Cfg.OFXDecimal)
    If dLBal = NODATE Then
        OFXTag "DTASOF", "00000000"
    Else
        OFXTag "DTASOF", FormatOFXDateTime(dLBal)
    End If
    OFXCloseTag "LEDGERBAL"
    If bABalPresent Then
        OFXOpenTag "AVAILBAL"
        OFXTag "BALAMT", FormatAmount(aABal, Cfg.OFXDecimal)
        OFXTag "DTASOF", FormatOFXDateTime(dABal)
        OFXCloseTag "AVAILBAL"
    End If
    Select Case sMyAcctType
    Case csCreditCard
        OFXCloseTag "CCSTMTRS"
        OFXCloseTag "CCSTMTTRNRS"
    Case Else
        OFXCloseTag "STMTRS"
        OFXCloseTag "STMTTRNRS"
    End Select
End Sub
Public Sub OFXFileTrailer(oFile As OutputFile)
    Set OutFile = oFile
    OFXCloseTag "OFX"
    Debug.Assert (Indent = 0)
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OFXSetVersion
' Description:       Set up for a certain OFX version
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       01/10/2004-14:09:58
'
' 20041201 CS Added 2.0.2 support
' 20080104 CS Added 2.1.0 support
' 20090417 CS Added 2.1.1 support
' Parameters :       Version (Integer) OFX Version. Must be 102, 151, 160, 200, 201, 202, 203, 210 or 211.
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub OFXSetVersion(Version As Integer)
    Select Case Version
    Case 102, 151, 160
        OFXHeader = 100
        OFXVersion = Version
    Case 200, 201, 202, 203, 210, 211
        OFXHeader = 200
        OFXVersion = Version
    Case Else
        LogMessage True, True, "Bad OFX Version value: " & CStr(Version) & ", defaulting to 102"
        OFXHeader = 100
        OFXVersion = 102
    End Select
End Sub
