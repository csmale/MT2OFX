Attribute VB_Name = "ProcessTransaction"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : FindPayeeFunc
'    Project    : MT2OFX
'
'    Description:
'
'    Modified   : $Author: Colin $ $Date: 14/11/10 23:58 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/FindPayee.bas 19    14/11/10 23:58 Colin $"
' $History: FindPayee.bas $
' 
' *****************  Version 19  *****************
' User: Colin        Date: 14/11/10   Time: 23:58
' Updated in $/MT2OFX
'
' *****************  Version 18  *****************
' User: Colin        Date: 25/11/08   Time: 22:20
' Updated in $/MT2OFX
' moving vss server!
'
' *****************  Version 16  *****************
' User: Colin        Date: 19/04/08   Time: 23:06
' Updated in $/MT2OFX
'
' *****************  Version 13  *****************
' User: Colin        Date: 2/11/05    Time: 23:03
' Updated in $/MT2OFX
' V3.4 beta 1
'
' *****************  Version 12  *****************
' User: Colin        Date: 18/03/05   Time: 21:57
' Updated in $/MT2OFX
'</CSCC>

Public Function FindServerTime(sFilein As String) As Date
    FindServerTime = GetFileModTime(sFilein)
End Function

Public Function DoProcessTxn(sBank As String, t As Txn) As Boolean
    Dim stmt As Statement
    Set stmt = t.Statement
    Dim sTmp As String
    Dim sStep As String
    Dim sHash As String
    
    sStep = "Initialisation"
    On Error GoTo baleout
    
    DoProcessTxn = False
    Dim oPayeeMapItem As PayeeMapItem
    
    If Bcfg.BankKey <> "" And Bcfg.ScriptFile <> "" Then
        sStep = "ScriptProcessTxn"
        If Not ScriptProcessTxn(sBank, t) Then
            Exit Function
        End If
    End If
    
    If Not Cfg.UseOldFITID Then
        sHash = t.GetHash   ' before payee mapping is applied!
    End If
    
    If Not Session.PayeeMap Is Nothing Then
        If Not t.SkipPayeeMapping Then
            sStep = "PayeeMapping"
            Session.PayeeMap.IgnoreCase = Cfg.PayeeMapIgnoreCase
            Set oPayeeMapItem = Session.PayeeMap.MapSearch(t.Memo, t.Payee, sTmp)
            If Not oPayeeMapItem Is Nothing Then
                t.Payee = sTmp
                If Len(oPayeeMapItem.Category) > 0 Then
                    t.Category = oPayeeMapItem.Category
                End If
                If oPayeeMapItem.SIC > 0 Then
                    t.SIC = oPayeeMapItem.SIC
                End If
            End If
        End If
    End If

' 7 Oct 2004: payee case stuff also governed by SkipPayeeMapping
' format payee string
    If Not t.SkipPayeeMapping Then
        sStep = "PayeeSetCase"
        t.Payee = PayeeSetCase(t.Payee, Cfg.PayeeCase)
        If Cfg.CompressSpaces Then
            sStep = "CompressSpaces"
            t.Payee = CompressSpaces(t.Payee)
            t.FurtherInfo = CompressSpaces(t.FurtherInfo)
        End If
    End If
    
' transaction type code
    If t.TxnType = "" Then
        sStep = "OFXTxnType"
        t.TxnType = OFXTxnType(t.BookingCode, t.Amt)
    End If
' assume OFX transaction type - must be translated for OFC
    If Session.FileFormat = FileFormatOFC Then
        sStep = "OFXTxnType2OFC"
        t.TxnType = OFXTxnType2OFC(t.TxnType)
    End If
' unique transaction ID - create from year, account, statement and a sequence.
    If t.FITID = "" Then
        sStep = "FITID"
        If Cfg.UseOldFITID Then
            If stmt.OFXStatementID = "" Then
                t.FITID = stmt.Acct & "." _
                    & CStr(Year(stmt.ClosingBalance.BalDate)) _
                    & Format(DatePart("y", stmt.ClosingBalance.BalDate), "000") _
                    & Format(t.Index, "000")
            Else ' no statement ID
    ' assumed that the statement numbers get reset at year end!
                t.FITID = CStr(Year(stmt.ClosingBalance.BalDate)) & "." _
                    & stmt.OFXStatementID & "." _
                    & Format(t.Index, "000")
            End If
        Else
            t.FITID = sHash
        End If
    End If
' cheque number - max 12 chars!
' CS 20040826: CHECKNUM is optional in OFX. Allow the script to set it if required
' and skip the defaulting code here. OFX/OFC Writer also changed to skip tag if empty
    If Cfg.GenerateCheckNum Then
        If t.CheckNum = "" Then
            sStep = "CheckNum"
            If stmt.StatementID = "" Or (stmt.StatementNum = 0 And stmt.StatementSeq = 1) Then
                t.CheckNum = Format(Year(stmt.ClosingBalance.BalDate), "0000") _
                    & Format(DatePart("y", stmt.ClosingBalance.BalDate), "000") _
                    & "/" & Format(t.Index, "000")
            Else
                t.CheckNum = Right$(stmt.StatementID, 8) & "/" & Format(t.Index, "000")
            End If
        End If
    End If
' CS 20041231: Default payee stuff never got finished somehow!
    If t.Payee = "" Then
        t.Payee = Cfg.DefaultPayee
    End If
    
    DoProcessTxn = True
goback:
    Exit Function
baleout:
    ShowError "DoProcessTxn", "Step=" & sStep
    Resume goback
End Function
