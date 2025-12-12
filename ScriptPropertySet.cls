VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScriptPropertySet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ScriptProperty"
Attribute VB_Ext_KEY = "Member0" ,"ScriptProperty"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : ScriptPropertySet
'    Project    : MT2OFX
'
'    Description:
'
'    Modified   : $Author: Colin $ $Date: 7/12/06 15:07 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/ScriptPropertySet.cls 9     7/12/06 15:07 Colin $"
' $History: ScriptPropertySet.cls $
' 
' *****************  Version 9  *****************
' User: Colin        Date: 7/12/06    Time: 15:07
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 6  *****************
' User: Colin        Date: 6/03/05    Time: 23:42
' Updated in $/MT2OFX
'</CSCC>

'local variable to hold collection
Private mCol As Collection

Public Function Add(objNewMember As ScriptProperty) As ScriptProperty
    mCol.Add objNewMember, objNewMember.Key

    'return the object created
    Set Add = objNewMember
End Function

Public Property Get Item(vntIndexKey As Variant) As ScriptProperty
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Load
' Description:       load the property set
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12/11/2004-22:02:43
'
' Parameters :       PropSetName: name of property set to load
'                    Props: array of arrays
'--------------------------------------------------------------------------------
'</CSCM>
Public Function Load(PropSetName As String, Props As Variant) As Boolean
    Set mCol = New Collection
    Dim sFile As String
' 20061129 CS: changed to app data for Vista compatibility
    sFile = Cfg.AppDataPath & "\sprops.ini"
    Dim sTmp As String
    Dim X As Variant    ' each x will be an array
    Dim xProp As ScriptProperty
    Load = False
    If IsEmpty(Props) Then
        Exit Function
    End If
    On Error Resume Next
    For Each X In Props
        Set xProp = New ScriptProperty
        xProp.Key = X(0)    ' key name
        xProp.Name = X(1)
        xProp.Description = X(2)
        xProp.DataType = X(3)   ' data type
        sTmp = ReadIniString(sFile, PropSetName, xProp.Key)
' other data types must be handled here!
        Select Case xProp.DataType
        Case ptBoolean
            Err.Clear
            xProp.Value = CBool(sTmp)
            If Err.Number > 0 Then
                Err.Clear
                xProp.Value = False
            End If
        Case ptDate
            If IsDate(sTmp) Then
                xProp.Value = CDate(sTmp)
            Else
                xProp.Value = Now()
            End If
        Case ptFloat
            If IsNumeric(sTmp) Then
                xProp.Value = CDbl(sTmp)
            Else
                xProp.Value = 0#
            End If
        Case ptInteger
            If IsNumeric(sTmp) Then
                xProp.Value = CLng(sTmp)
            Else
                xProp.Value = 0
            End If
        Case Else
            xProp.Value = sTmp
        End Select
        If xProp.DataType = ptChoice Then
            If UBound(X) > 3 Then
                If IsArray(X(4)) Then
                    xProp.Choices = X(4)
                Else
                    xProp.Choices = Array(X(4))
                End If
            Else
                xProp.Choices = Array()
            End If
        End If
        If UBound(X) > 4 Then
            xProp.Pattern = X(5)
        End If
        If UBound(X) > 5 Then
            xProp.ValidationFailMessage = X(6)
        End If
        Add xProp
    Next
    Load = True
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Save
' Description:       save the property set
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12/11/2004-22:03:45
'
' Parameters :       PropSetName (String)
'                    Props: array of arrays
'--------------------------------------------------------------------------------
'</CSCM>
Public Function Save(PropSetName As String, Props As Variant) As Boolean
    Dim sFile As String
' 20061129 CS: changed to app data for Vista compatibility
    sFile = Cfg.AppDataPath & "\sprops.ini"
    Dim X As Variant    ' each x will be a ScriptProperty
    Dim xProp As ScriptProperty
    For Each X In mCol
        Set xProp = X
' other data types must be handled here!
        WritePrivateProfileString PropSetName, xProp.Key, CStr(xProp.Value), sFile
    Next
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Find
' Description:       Search for and return a property, or Nothing if not found
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12/11/2004-22:14:25
'
' Parameters :       PropName (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function Find(PropName As String) As ScriptProperty
    On Error Resume Next
    Set Find = Nothing
    Set Find = mCol(PropName)
End Function
