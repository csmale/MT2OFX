VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Str86Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Payee"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : Str86Info
'    Project    : MT2OFX
'
'    Description:
'
'    Modified   : $Author: Colin $ $Date: 7/12/06 15:07 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/Str86Info.cls 6     7/12/06 15:07 Colin $"
' $History: Str86Info.cls $
' 
' *****************  Version 6  *****************
' User: Colin        Date: 7/12/06    Time: 15:07
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 3  *****************
' User: Colin        Date: 6/03/05    Time: 23:42
' Updated in $/MT2OFX
'</CSCC>

Public BusCode As Integer
Public Fields As New StructuredFields

Public Property Get Memo() As String
    Dim sTmp As String
    Dim sFld As String
    Dim i As Integer
    For i = sfDetails0 To sfDetailsLast
        sFld = Trim$(GetField(i))
        If (Not Bcfg.SkipEmptyMemoFields) Or sFld <> "" Then
            If sTmp = "" Then
                sTmp = sFld
            Else
                sTmp = sTmp & Cfg.MemoDelimiter & sFld
            End If
        End If
    Next
    Memo = sTmp
End Property

Public Property Get GetField(iFld As Integer) As String
    Dim i As Integer
    For i = 1 To Fields.Count
        If Fields(i).FieldNum = iFld Then
            GetField = Fields(i).Text
            Exit Property
        End If
    Next
    GetField = ""
End Property

Public Property Get Payee() As String
    Dim sPayee As String
    sPayee = GetField(sfNamePayee)
    If Right$(sPayee, 1) = " " Then
        sPayee = Trim$(sPayee) & " "
    End If
    Payee = Trim$(sPayee & Trim$(GetField(sfNamePayee2)))
End Property

Public Function ParseInfo(sText As String) As Boolean
    ParseInfo = False
    BusCode = -1
    Dim i As Integer
    Set Fields = New StructuredFields
    If Not IsNumeric(Left$(sText, 3)) Then Exit Function
    BusCode = Val(Left$(sText, 3))
    Dim sDelim As String
    sDelim = Mid$(sText, 4, 1)
    Dim iPtr As Integer
    iPtr = 5
    Dim sTmp As String
    While Len(sText) > iPtr
        i = InStr(iPtr, sText, sDelim)
        If i = 0 Then
            sTmp = Mid$(sText, iPtr)
            iPtr = 9999
        Else
            sTmp = Mid$(sText, iPtr, i - iPtr)
            iPtr = i + 1
        End If
        Fields.Add Mid$(sTmp, 3), Val(Left$(sTmp, 2))
    Wend
    ParseInfo = True
End Function

