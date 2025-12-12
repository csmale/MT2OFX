VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Statement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Txns"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"Session"
Attribute VB_Ext_KEY = "Member2" ,"Session"
Attribute VB_Ext_KEY = "Member3" ,"Txns"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : Statement
'    Project    : MT2OFX
'
'    Description:
'
'    Modified   : $Author: Colin $ $Date: 7/12/06 15:07 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/Statement.cls 11    7/12/06 15:07 Colin $"
' $History: Statement.cls $
' 
' *****************  Version 11  *****************
' User: Colin        Date: 7/12/06    Time: 15:07
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 10  *****************
' User: Colin        Date: 25/04/06   Time: 21:45
' Updated in $/MT2OFX
'
' *****************  Version 7  *****************
' User: Colin        Date: 2/11/05    Time: 23:03
' Updated in $/MT2OFX
' V3.4 beta 1
'
' *****************  Version 6  *****************
' User: Colin        Date: 6/03/05    Time: 23:42
' Updated in $/MT2OFX
'</CSCC>

Public Session As Session

Public Acct As String
' 20050905 CS: Added account type to Statement. Should be "CHECKING", "CREDITLINE", "MONEYMRKT", "SAVINGS"
' as per OFX specs. Can also be "CREDITCARD" - which triggers special creditcard OFX processing.
Public AcctType As String
' 20051014 CS: Added QIFAcctType to override the "!Type:" header
Public QIFAcctType As String
Public BankName As String
' 20041208 CS Added Branch Name to support CyberGestion
Public BranchName As String
Private xStatementID As String
Public RelatedReference As String
Public OpeningBalance As New Balance
Public ClosingBalance As New Balance
Public AvailableBalance As New Balance
Private xNonSwift As New NonSwiftInfo
Private xTxns As New Txns
Public Index As Long
Public StatementNum As Integer
Public StatementSeq As Integer
Public OFXStatementID As String
Private xSecurityList As New SecurityList

Private Sub Class_Initialize()
    Acct = ""
    BankName = ""
    BranchName = ""
    RelatedReference = ""
    xStatementID = ""
    StatementNum = 0
    StatementSeq = 0
    Index = -1
    OFXStatementID = ""
End Sub

Public Property Get Txns() As Txns
    Set Txns = xTxns
End Property

Public Property Get SecurityList() As SecurityList
    Set SecurityList = xSecurityList
End Property

Public Property Get NonSwift() As NonSwiftInfo
    Set NonSwift = xNonSwift
End Property

Public Property Get StatementID() As String
    StatementID = xStatementID
End Property

Public Property Get TotalCredits() As Double
    Dim t As Txn
    Dim dTot As Double
    dTot = 0
    For Each t In xTxns
        If t.Amt >= 0 Then dTot = dTot + t.Amt
    Next
    TotalCredits = dTot
End Property

Public Property Get TotalDebits() As Double
    Dim t As Txn
    Dim dTot As Double
    dTot = 0
    For Each t In xTxns
        If t.Amt < 0 Then dTot = dTot - t.Amt
    Next
    TotalDebits = dTot
End Property

Public Property Get CountCredits() As Long
    Dim t As Txn
    Dim iCnt As Long
    iCnt = 0
    For Each t In xTxns
        If t.Amt >= 0 Then iCnt = iCnt + 1
    Next
    CountCredits = iCnt
End Property

Public Property Get CountDebits() As Long
    Dim t As Txn
    Dim iCnt As Long
    iCnt = 0
    For Each t In xTxns
        If t.Amt < 0 Then iCnt = iCnt + 1
    Next
    CountDebits = iCnt
End Property

Public Property Let StatementID(sNew As String)
    Dim iSlash As Integer
    xStatementID = sNew
    On Error GoTo baleout
    StatementNum = Val(sNew)
    iSlash = InStr(sNew, "/")
    If iSlash = 0 Then
        StatementSeq = 0
    Else
        StatementSeq = Val(Mid$(sNew, iSlash + 1))
    End If
    Exit Property
baleout:
    StatementNum = 0
    StatementSeq = 0
End Property
