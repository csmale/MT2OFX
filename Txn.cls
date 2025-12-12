VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Txn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Splits"
Attribute VB_Ext_KEY = "Member1" ,"Statement"
Attribute VB_Ext_KEY = "Member2" ,"Payee"
Attribute VB_Ext_KEY = "Member3" ,"Splits"
Attribute VB_Ext_KEY = "Member4" ,"Statement"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : Txn
'    Project    : MT2OFX
'
'    Description:
'
'    Modified   : $Author: Colin $ $Date: 14/11/10 23:51 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/Txn.cls 18    14/11/10 23:51 Colin $"
' $History: Txn.cls $
' 
' *****************  Version 18  *****************
' User: Colin        Date: 14/11/10   Time: 23:51
' Updated in $/MT2OFX
'
' *****************  Version 17  *****************
' User: Colin        Date: 25/11/08   Time: 22:24
' Updated in $/MT2OFX
' moving vss server!
'
' *****************  Version 15  *****************
' User: Colin        Date: 19/04/08   Time: 22:13
' Updated in $/MT2OFX
' added SupplementaryDetails
'
' *****************  Version 14  *****************
' User: Colin        Date: 7/12/06    Time: 15:07
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 11  *****************
' User: Colin        Date: 2/11/05    Time: 23:03
' Updated in $/MT2OFX
' V3.4 beta 1
'
' *****************  Version 10  *****************
' User: Colin        Date: 6/03/05    Time: 23:42
' Updated in $/MT2OFX
'</CSCC>

Private xStatement As Statement
Private xIndex As Integer
Public Amt As Double
Public ValueDate As Date
Public BookDate As Date
Public IsReversal As Boolean
Public BookingCode As String
Public Reference As String
Public BankReference As String
Public SupplementaryDetails As String
Public FurtherInfo As String
Private xStr86 As Str86Info
Private xNonSwift As NonSwiftInfo

' the following are derived indirectly from data contained in the
' MT940 file, normally by parsing the description in the :86: record(s)
' and the txn type in the :61: record
' 20051013 CS: Payee is now an object
Public Payee As New Payee
Public TxnDateValid As Boolean
Public TxnDate As Date
Public TxnType As String    ' must be valid OFX in <TRANTYPE>
Public FITID As String
Public CheckNum As String
' 11 May 2004 - Added Category for use with QIF
Public Category As String
' 21 July 2004 - Added SkipPayeeMapping
Public SkipPayeeMapping As Boolean
' 12 August 2004 - Added ClearedStatus
Public ClearedStatus As String
' 6 October 2004 - Split Support
Private xSplits As Splits
' 20050918 CS: Added SIC code for txn
Public SIC As Long
' 20051013 CS: Added CustomQIFItems
Public CustomQIFItems As New Scripting.Dictionary

Public Sub Clear()
    Str86.BusCode = -1
    Reference = ""
    FurtherInfo = ""
    SupplementaryDetails = ""
    TxnType = ""
    Category = ""
    TxnDateValid = False
    SkipPayeeMapping = False
    ClearedStatus = ""
    ValueDate = NODATE
    TxnDate = NODATE
    BookDate = NODATE
End Sub

Public Property Get Statement() As Statement
    Set Statement = xStatement
End Property
Friend Property Set Statement(newStatement As Statement)
    Set xStatement = newStatement
End Property
Public Property Get Index() As Integer
    Index = xIndex
End Property
Friend Property Let Index(newIndex As Integer)
    xIndex = newIndex
End Property
Public Property Get Str86() As Str86Info
    Set Str86 = xStr86
End Property
Public Property Get NonSwift() As NonSwiftInfo
    Set NonSwift = xNonSwift
End Property
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Splits
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       06/10/2004-21:43:12
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get Splits() As Splits
    Set Splits = xSplits
End Property
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Memo
' Description:       Alias of FurtherInfo
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       29/09/2004-23:03:24
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get Memo() As String
    Memo = FurtherInfo
End Property
Public Property Let Memo(sNewMemo As String)
    FurtherInfo = sNewMemo
End Property


'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetHash
' Description:       Returns an MD5 hash of a selection of the transaction fields, as a string (32 hex chars)
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       5/11/2008-21:35:08
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetHash() As String
    Dim sTmp As String
    Dim sInput As String
    Dim bOut() As Byte

' select, format and concatenate the fields for the hash
' as this is used to create a FITID, the fields must be sufficient to guarantee uniqueness within the
' account.
' all formats must be locale-neutral.
' use of Payee and Memo are controversial because they can be modified by Payee Mapping.
' it would be best to establish the transaction hash BEFORE the payee mapping is applied.
    sInput = Memo _
            & Payee.Name _
            & Format(BookDate, "yyyymmdd") _
            & FormatAmount(Amt, ".")

    If MD5HashStart() Then
        If MD5HashDataString(sInput) Then
            If MD5HashEnd(bOut) Then
                sTmp = Base64Encode(bOut)
'                sTmp = HashHexString(bOut)
            Else
                sTmp = HashErrorString("Error from MD5HashEnd")
            End If
        Else
            sTmp = HashErrorString("Error from MD5HashDataString")
        End If
    Else
        sTmp = HashErrorString("Error from MD5HashStart")
    End If
    
    GetHash = sTmp
End Function

Private Sub Class_Initialize()
    Set xNonSwift = New NonSwiftInfo
    Set xStr86 = New Str86Info
    Set xSplits = New Splits
    Clear
End Sub

