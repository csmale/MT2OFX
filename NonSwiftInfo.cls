VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NonSwiftInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : NonSwiftInfo
'    Project    : MT2OFX
'
'    Description:
'
'    Modified   : $Author: Colin $ $Date: 7/12/06 15:07 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/NonSwiftInfo.cls 7     7/12/06 15:07 Colin $"
' $History: NonSwiftInfo.cls $
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

Public Fields As New StructuredFields

Public Function ParseInfo(sText As String) As Boolean
    ParseInfo = False
    Set Fields = New StructuredFields
    Dim i As Integer
    Dim vArr As Variant
    vArr = Split(sText, vbLf)
    Dim sTmp As String
    For i = LBound(vArr) To UBound(vArr)
        sTmp = vArr(i)
        Fields.Add Trim$(Mid$(sTmp, 3)), Val(Left$(sTmp, 2))
    Next
    ParseInfo = True
End Function

Public Property Get GetField(iFld As Integer) As String
    Dim i As Integer
    For i = 1 To Fields.Count
        If Fields(i).FieldNum = iFld Then
            GetField = Fields(i).Text
'               Debug.Print "NS(" & iFld & ") = " & GetField
            Exit Property
        End If
    Next
    GetField = ""
End Property

