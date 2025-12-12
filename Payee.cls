VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Payee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : Payee
'    Project    : MT2OFX
'
'    Description: Payee Class, modelled on OFX but with extra AddrN fields for QIF
'
'    Modified   : $Author: Colin $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/Payee.cls 7     6/10/09 0:21 Colin $"
' $History: Payee.cls $
' 
' *****************  Version 7  *****************
' User: Colin        Date: 6/10/09    Time: 0:21
' Updated in $/MT2OFX
' added payee account fields
'
' *****************  Version 5  *****************
' User: Colin        Date: 7/12/06    Time: 15:07
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 2  *****************
' User: Colin        Date: 13/10/05   Time: 23:42
' Updated in $/MT2OFX

'</CSCC>

Public Name As String
Attribute Name.VB_VarUserMemId = 0
Private xAddr(1 To 6) As String
' NB: Addr4..6 are not used in OFX
' NB: City, State, PostalCode, Country and Phone are not used in QIF
Public City As String
Public State As String
Public PostalCode As String
Public Country As String
Public Phone As String
' 20091005 CS: Add payee bank account info
' for more details see Statement
Public Acct As String
Public AcctType As String
Public BankName As String
Public BranchName As String
Public AcctKey As String


'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       IsSimple
' Description:       Determines if a Payee can be represented by the name field alone
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/13/2005-12:44:07
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function IsSimple() As Boolean
    IsSimple = (Len(xAddr(1)) + Len(xAddr(2)) + Len(xAddr(3)) + Len(xAddr(4)) + Len(xAddr(5)) + Len(xAddr(6)) _
        + Len(City) + Len(State) + Len(PostalCode) _
        + Len(Country) + Len(Phone)) = 0
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       LastUsedAddressLine
' Description:       Returns index of last non-blank entry in the address array
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/13/2005-17:20:18
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function LastUsedAddressLine() As Integer
    Dim i As Integer
    For i = 6 To 1 Step -1
        If Len(Addr(i)) > 0 Then
            LastUsedAddressLine = i
            Exit Function
        End If
    Next
    LastUsedAddressLine = 0
End Function

Public Property Get Addr() As Variant
    Addr = xAddr
End Property

Public Property Get Addr1() As String
    Addr1 = xAddr(1)
End Property
Public Property Let Addr1(sAddr As String)
    xAddr(1) = sAddr
End Property
Public Property Get Addr2() As String
    Addr2 = xAddr(2)
End Property
Public Property Let Addr2(sAddr As String)
    xAddr(2) = sAddr
End Property
Public Property Get Addr3() As String
    Addr3 = xAddr(3)
End Property
Public Property Let Addr3(sAddr As String)
    xAddr(3) = sAddr
End Property
Public Property Get Addr4() As String
    Addr3 = xAddr(4)
End Property
Public Property Let Addr4(sAddr As String)
    xAddr(4) = sAddr
End Property
Public Property Get Addr5() As String
    Addr3 = xAddr(5)
End Property
Public Property Let Addr5(sAddr As String)
    xAddr(5) = sAddr
End Property
Public Property Get Addr6() As String
    Addr3 = xAddr(6)
End Property
Public Property Let Addr6(sAddr As String)
    xAddr(6) = sAddr
End Property
