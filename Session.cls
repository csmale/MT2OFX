VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Statements"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"PayeeMap"
Attribute VB_Ext_KEY = "Member2" ,"OutputFile"
Attribute VB_Ext_KEY = "Member3" ,"InputFile"
Attribute VB_Ext_KEY = "Member4" ,"OutputFile"
Attribute VB_Ext_KEY = "Member5" ,"InputFile"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : Session
'    Project    : MT2OFX
'
'    Description:
'
'    Modified   : $Author: Colin $ $Date: 15/11/10 0:06 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/Session.cls 18    15/11/10 0:06 Colin $"
' $History: Session.cls $
' 
' *****************  Version 18  *****************
' User: Colin        Date: 15/11/10   Time: 0:06
' Updated in $/MT2OFX
'
' *****************  Version 17  *****************
' User: Colin        Date: 25/11/08   Time: 22:24
' Updated in $/MT2OFX
' moving vss server!
'
' *****************  Version 15  *****************
' User: Colin        Date: 20/04/08   Time: 10:06
' Updated in $/MT2OFX
' For 3.5 beta 1
'
' *****************  Version 14  *****************
' User: Colin        Date: 7/12/06    Time: 15:07
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 13  *****************
' User: Colin        Date: 25/04/06   Time: 21:45
' Updated in $/MT2OFX
'
' *****************  Version 10  *****************
' User: Colin        Date: 6/03/05    Time: 23:42
' Updated in $/MT2OFX
'</CSCC>

Public FileIn As String
Public InputFile As New InputFile
Public FileOut As String
Public OutputFile As New OutputFile

' 20060330 CS: Structure output file format, separate from actual file name
Public Enum zFileFormat
    FileFormatUnknown = 0
    FileFormatQIF = 1
    FileFormatOFX = 2
    FileFormatQFX = 3
    FileFormatOFC = 4
' 20100423 CS: custom output formats
    FileFormatCustom = 5
End Enum
Public FileFormat As zFileFormat
Public CustomFormatName As String

Rem Public InputHandler As String
Rem Public OutputHandler As String

Public BankID As String
'Public OutputFileType As String
Public ServerTime As Date
Public PayeeMapIgnoreCase As Boolean
' 20050220 CS: Add language code to session
Public Language As String

Private xPayeeMapFile As String
Private xPayeeMap As PayeeMap
Private xStmts As Statements

Public SuppressedStatementCount As Long  ' number of suppressed statements

Public Property Get OutputFileType() As String
    Dim sRet As String
    Select Case FileFormat
    Case FileFormatQIF
        sRet = "QIF"
    Case FileFormatOFX
        sRet = "OFX"
    Case FileFormatQFX
        sRet = "QFX"
    Case FileFormatOFC
        sRet = "OFC"
    Case FileFormatCustom
        sRet = CustomFormatName
    Case Else
        sRet = ""
    End Select
    OutputFileType = sRet
End Property

Friend Property Let OutputFileType(sType As String)
    Select Case UCase$(sType)
    Case "QIF"
        FileFormat = FileFormatQIF
'        OutputHandler = ohQIF
    Case "OFX"
        FileFormat = FileFormatOFX
'        OutputHandler = ohOFX
    Case "QFX"
        FileFormat = FileFormatQFX
'        OutputHandler = ohQFX
    Case "OFC"
        FileFormat = FileFormatOFC
'        OutputHandler = ohOFC
    Case ""
        FileFormat = FileFormatUnknown
'        OutputHandler = ohUnknown
    Case Else
        FileFormat = FileFormatCustom
'        OutputHandler = sType
    End Select
End Property

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       PayeeMap
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       10/02/2005-12:06:07
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get PayeeMap() As PayeeMap
    Set PayeeMap = xPayeeMap
End Property

Public Property Get PayeeMapFile() As String
    PayeeMapFile = xPayeeMapFile
End Property

Public Property Let PayeeMapFile(sFile As String)
    Set xPayeeMap = New PayeeMap
    If Not xPayeeMap.LoadPayeeMap(sFile) Then
        xPayeeMapFile = ""
        Set xPayeeMap = Nothing
    Else
        xPayeeMapFile = sFile
        LogMessage False, True, " Payee Map set to: " & xPayeeMapFile
    End If
End Property

Public Property Get Statements() As Statements
    Set Statements = xStmts
End Property

Public Property Get TransactionCount() As Long
    Dim s As Statement
    Dim iCnt As Long
    iCnt = 0
    For Each s In xStmts
        iCnt = iCnt + s.Txns.Count
    Next
    TransactionCount = iCnt
End Property

Public Property Get CountCredits() As Long
    Dim s As Statement
    Dim iCnt As Long
    iCnt = 0
    For Each s In xStmts
        iCnt = iCnt + s.CountCredits
    Next
    CountCredits = iCnt
End Property

Public Property Get CountDebits() As Long
    Dim s As Statement
    Dim iCnt As Long
    iCnt = 0
    For Each s In xStmts
        iCnt = iCnt + s.CountDebits
    Next
    CountDebits = iCnt
End Property

Public Property Get TotalCredits() As Double
    Dim s As Statement
    Dim dTot As Double
    dTot = 0
    For Each s In xStmts
        dTot = dTot + s.TotalCredits
    Next
    TotalCredits = dTot
End Property

Public Property Get TotalDebits() As Double
    Dim s As Statement
    Dim dTot As Double
    dTot = 0
    For Each s In xStmts
        dTot = dTot + s.TotalDebits
    Next
    TotalDebits = dTot
End Property

Private Sub Class_Initialize()
    Set xStmts = New Statements
    FileFormat = FileFormatUnknown
    ServerTime = NODATE
End Sub
