VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScriptProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : ScriptProperty
'    Project    : MT2OFX
'
'    Description:
'
'    Modified   : $Author: Colin $ $Date: 7/12/06 15:07 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/ScriptProperty.cls 7     7/12/06 15:07 Colin $"
' $History: ScriptProperty.cls $
' 
' *****************  Version 7  *****************
' User: Colin        Date: 7/12/06    Time: 15:07
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 4  *****************
' User: Colin        Date: 6/03/05    Time: 23:42
' Updated in $/MT2OFX
'</CSCC>

Public Enum PropDataType
    ptString = 1
    ptInteger = 2
    ptDate = 3
    ptBoolean = 4
    ptFloat = 5
    ptChoice = 6
    ptCurrency = 7
End Enum

Public Key As String
Public DataType As PropDataType
Public Value As Variant
Public Name As String
Public Description As String
Public Choices As Variant   ' contains string array!
Public Pattern As String    ' regexp pattern for validation
Public ValidationFailMessage As String  ' message given on validation failure


Private Sub Class_Initialize()
    Choices = Array()
End Sub
