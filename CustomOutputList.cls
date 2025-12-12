VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CustomOutputList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CustomOutputFormat"
Attribute VB_Ext_KEY = "Member0" ,"CustomOutputFormat"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : CustomOutputList
'    Project    : MT2OFX
'
'    Description: [type_description_here]
'
'    Modified   : $Author: Colin $ $Date: 15/11/10 0:24 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/CustomOutputList.cls 2     15/11/10 0:24 Colin $"
' $History: CustomOutputList.cls $
' 
' *****************  Version 2  *****************
' User: Colin        Date: 15/11/10   Time: 0:24
' Updated in $/MT2OFX
'</CSCC>
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function Add(objNewMember As CustomOutputFormat) As CustomOutputFormat
    mCol.Add objNewMember, objNewMember.FormatName

    'return the object created
    Set Add = objNewMember
End Function

Public Property Get Item(vntIndexKey As Variant) As CustomOutputFormat
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
' Description:       [type_description_here]
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       21/04/2010-23:11:29
'
' Parameters :       sDir (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function Load(sDir As String) As Boolean
    Dim xNew As CustomOutputFormat
    Dim fso As FileSystemObject
    Dim fDir As Scripting.Folder
    Dim f As Scripting.File
    Dim sLang As String
    Dim sMod As String
    Dim sScript As String
    Dim bRet As Boolean
    
    Load = False
    If Not InitialiseScripting() Then
        Exit Function
    End If

    On Error GoTo baleout
    Set fso = New Scripting.FileSystemObject
    Set mCol = New Collection
    Set fDir = fso.GetFolder(sDir)
    For Each f In fDir.Files
        If LCase(GetExtension(f.Name)) <> "vbs" Then GoTo trynext
        sScript = f.Path
        sMod = ScriptModuleName(f.Name)
        LogMessage False, True, "Loading output script file: " & sScript, ""
        GetScriptEnv().AbortRequested = False
        If Not ScriptInit(sMod, sScript) Then
' message 118 is "Processing aborted" - we must be able to improve on this!
            LogMessage True, True, GetString(118), AppName
            GoTo trynext
        End If
' call ScriptConfig in this script
        Set xNew = New CustomOutputFormat
        xNew.ScriptName = f.Name
        If CallScriptFunction(sMod, "ScriptConfig", xNew, bRet) Then
            If bRet Then
                Add xNew
            End If
        End If
trynext:
    Next
    Load = True
    Exit Function
baleout:
    Load = False
    Exit Function
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       FindByFormatName
' Description:       Find a CustomOutputFormat by name in the collection
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       08/05/2010-15:44:51
'
' Parameters :       sName (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function FindByFormatName(sName As String) As CustomOutputFormat
    Dim xRet As CustomOutputFormat
    Dim sNameL As String
    sNameL = LCase$(sName)
    For Each xRet In mCol
        If LCase(xRet.FormatName) = sNameL Then
            Set FindByFormatName = xRet
            Exit Function
        End If
    Next
    Set FindByFormatName = Nothing
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       FindByScriptName
' Description:       Find a CustomOutputFormat by Script Name
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       08/05/2010-15:47:26
'
' Parameters :       sName (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function FindByScriptName(sName As String) As CustomOutputFormat
    Dim xRet As CustomOutputFormat
    Dim sNameL As String
    sNameL = LCase$(sName)
    For Each xRet In mCol
        If LCase(xRet.ScriptName) = sNameL Then
            Set FindByScriptName = xRet
            Exit Function
        End If
    Next
    Set FindByScriptName = Nothing
End Function
