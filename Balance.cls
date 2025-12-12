VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Balance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : Balance
'    Project    : MT2OFX
'
'    Description: Balance Class
'
'    Modified   : $Author: Colin $ $Date: 7/12/06 14:53 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/Balance.cls 9     7/12/06 14:53 Colin $"
Rem $History: Balance.cls $
' 
' *****************  Version 9  *****************
' User: Colin        Date: 7/12/06    Time: 14:53
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 6  *****************
' User: Colin        Date: 18/04/05   Time: 21:28
' Updated in $/MT2OFX
' Add initialisation of BalDate
'
' *****************  Version 5  *****************
' User: Colin        Date: 6/03/05    Time: 23:41
' Updated in $/MT2OFX
'
' *****************  Version 4  *****************
' User: Colin        Date: 6/03/05    Time: 0:35
' Updated in $/MT2OFX
' *****************  Version 3  *****************
' User: Colin        Date: 6/03/05    Time: 0:17
' Updated in $/MT2OFX
'</CSCC>

Public Amt As Double
Public Ccy As String
Public BalDate As Date

Private Sub Class_Initialize()
    Amt = 0
    Ccy = ""
' 20050418 CS: Initialise balance date
    BalDate = NODATE
End Sub
