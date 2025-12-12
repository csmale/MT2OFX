Attribute VB_Name = "OFCWrite"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : OFCWrite
'    Project    : MT2OFX
'
'    Description: OFC Formatting
'
'    Modified   : $Author: Colin $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/OFCWrite.bas 11    20/04/08 10:06 Colin $"
' $History: OFCWrite.bas $
' 
' *****************  Version 11  *****************
' User: Colin        Date: 20/04/08   Time: 10:06
' Updated in $/MT2OFX
' For 3.5 beta 1

'</CSCC>

Private FileNum As Integer
Private FileOut As OutputFile
Private Indent As Integer
Private sIndent As String
Private SpacesPerIndent As Integer

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       FormatOFCDate
' Description:       Format a date for OFC
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/20/2007-11:02:58
'
' Parameters :       dDate (Date)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function FormatOFCDate(dDate As Date) As String
    Dim dTmp As Date
    dTmp = dDate
' if there is a non-zero time part, this time is in GMT and needs to be corrected to the local timezone
' in case this causes the date to change to the next/previous day!
    If (Hour(dTmp) + Minute(dTmp) + Second(dTmp)) > 0 Then
        dTmp = DateAdd("n", -GetCurrentTimeBias(), dTmp)
    End If
    FormatOFCDate = Format(dTmp, "yyyymmdd")
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OFCEscape
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/20/2007-11:02:58
'
' Parameters :       sLine (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function OFCEscape(sLine As String) As String
    Dim sTmp As String
    sTmp = Replace(sLine, "<", " ")
    sTmp = Replace(sTmp, ">", " ")
    sTmp = Replace(sTmp, "}", "")
    sTmp = Replace(sTmp, "{", "")
    OFCEscape = sTmp
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OFCPrintLine
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/20/2007-11:02:58
'
' Parameters :       sLine (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub OFCPrintLine(sLine As String)
    FileOut.PrintLine sLine
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OFCOpenTag
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/20/2007-11:02:58
'
' Parameters :       sTag (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub OFCOpenTag(sTag As String)
    OFCPrintLine sIndent & "<" & sTag & ">"
    Indent = Indent + 1
    sIndent = Space$(Indent * SpacesPerIndent)
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OFCTag
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/20/2007-11:02:58
'
' Parameters :       sTag (String)
'                    sText (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub OFCTag(sTag As String, sText As String)
    OFCPrintLine sIndent & "<" & sTag & ">" & OFCEscape(sText)
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OFCCloseTag
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/20/2007-11:02:58
'
' Parameters :       sTag (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub OFCCloseTag(sTag As String)
    Indent = Indent - 1
    sIndent = Space$(Indent * SpacesPerIndent)
    OFCPrintLine sIndent & "</" & sTag & ">"
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OFCFileHeader
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/20/2007-11:02:58
'
' Parameters :       oFile (OutputFile)
'                    dNow (Date)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub OFCFileHeader(oFile As OutputFile, dNow As Date)
    Set FileOut = oFile
    Indent = 0
    sIndent = ""
    SpacesPerIndent = 2
    OFCOpenTag "OFC"
    OFCTag "DTD", "2"
    OFCTag "CPAGE", "1252"
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OFCStatementHeader
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/20/2007-11:02:58
'
' Parameters :       oFile (OutputFile)
'                    sDefCur (String)
'                    sBankName (String)
'                    sBranch (String)
'                    sAcctNum (String)
'                    sAcctType (String)
'                    dFrom (Date)
'                    dTo (Date)
'                    aLBal (Double)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub OFCStatementHeader(oFile As OutputFile, _
    sDefCur As String, sBankName As String, sBranch As String, sAcctNum As String, _
    sAcctType As String, dFrom As Date, dTo As Date, aLBal As Double)
    Set FileOut = oFile
    OFCOpenTag "ACCTSTMT"
    OFCOpenTag "ACCTFROM"
    OFCTag "BANKID", sBankName
    If Len(sBranch) > 0 Then
        OFCTag "BRANCHID", sBranch
    End If
    OFCTag "ACCTID", sAcctNum
    OFCTag "ACCTTYPE", sAcctType
    OFCCloseTag "ACCTFROM"
    OFCOpenTag "STMTRS"
    OFCTag "DTSTART", FormatOFCDate(dFrom)
    OFCTag "DTEND", FormatOFCDate(dTo)
    Dim sAmt As String
    sAmt = Format(aLBal, "0.00")
    sAmt = Replace(sAmt, SystemDecimalSeparator(), ".")
    OFCTag "LEDGER", sAmt
End Sub
' 20050918 CS: Add SIC to signature
' 20051013 CS: Payee is now an object
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OFCTransaction
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/20/2007-11:02:58
'
' Parameters :       oFile (OutputFile)
'                    sTrnType (String)
'                    dValDate (Date)
'                    dTrnDate (Date)
'                    dBookDate (Date)
'                    bUseTrnDate (Boolean)
'                    aAmount (Double)
'                    sFITID (String)
'                    sCheckNum (String)
'                    xPayee (Payee)
'                    sMemo (String)
'                    iSIC (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub OFCTransaction(oFile As OutputFile, sTrnType As String, _
    dValDate As Date, dTrnDate As Date, dBookDate As Date, _
    bUseTrnDate As Boolean, aAmount As Double, sFITID As String, _
    sCheckNum As String, xPayee As Payee, sMemo As String, iSIC As Long)
    Set FileOut = oFile
    Dim sAmt As String
    sAmt = Format(aAmount, "0.00")
    sAmt = Replace(sAmt, SystemDecimalSeparator(), ".")
    OFCOpenTag "STMTTRN"
    OFCTag "TRNTYPE", sTrnType
    OFCTag "DTPOSTED", FormatOFCDate(dBookDate)
    OFCTag "TRNAMT", sAmt
    OFCTag "FITID", sFITID
' CS 20040826: Allow CHKNUM to be optional
    If Len(sCheckNum) > 0 Then
        OFCTag "CHKNUM", sCheckNum
    End If
' 20050918 CS: Add SIC
    If iSIC > 0 Then
        OFCTag "SIC", CStr(iSIC)
    End If
' 20051013 CS: Payee is now an object
    OFCPayee xPayee
    If Len(sMemo) > 0 Then
        OFCTag "MEMO", sMemo
    End If
    OFCCloseTag "STMTTRN"
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OFCPayee
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/20/2007-11:02:58
'
' Parameters :       xPayee (Payee)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub OFCPayee(xPayee As Payee)
    If xPayee.IsSimple Then
' CS 20050109: Allow NAME to be optional
        If Len(xPayee.Name) > 0 Then
            OFCTag "NAME", OFCString(xPayee.Name, 32)
        End If
    Else
        OFCOpenTag "PAYEE"
        OFCTag "NAME", OFCString(xPayee.Name, 32)
        OFCTag "ADDRESS", OFCString(xPayee.Addr1, 32)
        OFCTag "ADDRESS", OFCString(xPayee.Addr2, 32)
        OFCTag "CITY", OFCString(xPayee.City, 20)
        OFCTag "STATE", OFCString(xPayee.State, 2)
        OFCTag "POSTALID", OFCString(xPayee.PostalCode, 9)
        OFCTag "PHONE", OFCString(xPayee.Phone, 10)
        OFCCloseTag "PAYEE"
    End If
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OFCString
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/20/2007-11:02:58
'
' Parameters :       s (String)
'                    iMaxLen (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Private Function OFCString(s As String, iMaxLen As Long) As String
    Dim sTmp As String
    sTmp = Left$(Trim$(s), iMaxLen)
    If Len(sTmp) = 0 Then sTmp = "-"
    OFCString = sTmp
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OFCStatementTrailer
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/20/2007-11:02:58
'
' Parameters :       oFile (OutputFile)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub OFCStatementTrailer(oFile As OutputFile)
    Set FileOut = oFile
    OFCCloseTag "STMTRS"
    OFCCloseTag "ACCTSTMT"
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OFCFileTrailer
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/20/2007-11:02:58
'
' Parameters :       oFile (OutputFile)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub OFCFileTrailer(oFile As OutputFile)
    Set FileOut = oFile
    OFCCloseTag "OFC"
    Debug.Assert (Indent = 0)
End Sub

