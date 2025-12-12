VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PayeeMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"PayeeMap"
Attribute VB_Ext_KEY = "Member0" ,"PayeeMap"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : PayeeMap
'    Project    : MT2OFX
'
'    Description:
'
'    Modified   : $Author: Colin $ $Date: 7/12/06 15:07 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/PayeeMaps.cls 11    7/12/06 15:07 Colin $"
' $History: PayeeMaps.cls $
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

Public Delimiter As String
Public IgnoreCase As Boolean

'local variable to hold collection
Private mCol As Collection

Public Function Add(objNewMember As PayeeMapItem, Optional sKey As String) As PayeeMapItem
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
End Function

Public Property Get Item(vntIndexKey As Variant) As PayeeMapItem
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
    IgnoreCase = False
    Delimiter = SystemListSeparator()
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Function LoadPayeeMap(sFile As String) As Boolean
    Dim vFields As Variant
    Dim iFile As Integer
    Dim oPayee As PayeeMapItem
    Dim sLine As String
    Dim iTmp As Long
    
    LoadPayeeMap = False
    On Error GoTo nofile
    iFile = FreeFile
    Open sFile For Input Access Read As iFile
    On Error GoTo 0
    Do While Not EOF(iFile)
        Line Input #iFile, sLine
        sLine = Trim(sLine)
        If Left(sLine, 1) <> "#" Then
            vFields = txtParseLineDelimited(sLine, Delimiter, False)
            If TypeName(vFields) = "Variant()" Then
                If UBound(vFields) >= 1 Then
                    Set oPayee = New PayeeMapItem
                    oPayee.Pattern = vFields(1)
                    If UBound(vFields) > 1 Then
                        oPayee.Replacement = vFields(2)
                    End If
                    If UBound(vFields) > 2 Then
                        oPayee.Category = vFields(3)
                    End If
                    If UBound(vFields) > 3 Then
                        If IsNumeric(vFields(4)) Then
                            iTmp = CLng(vFields(4))
                            If iTmp = mfMemo Or iTmp = mfPayee Then
                                oPayee.MatchField = iTmp
                            Else
                                oPayee.MatchField = mfPayee
                            End If
                        Else
                            oPayee.MatchField = mfPayee
                        End If
                        If UBound(vFields) > 4 Then
                            oPayee.SIC = Val(vFields(5))
                            If UBound(vFields) > 5 Then
                                If IsNumeric(vFields(6)) Then
                                    oPayee.NoReturn = (Val(vFields(6)) > 0)
                                Else
                                    oPayee.NoReturn = False
                                End If
                            Else
                                oPayee.NoReturn = False
                            End If
                        End If
                    Else
                        oPayee.MatchField = mfPayee
                    End If
                    Add oPayee
                End If
            End If
        End If
    Loop
    Close #iFile
    LoadPayeeMap = True
    Exit Function
nofile:
' 53: file not found
' 76: path not found
    If Err.Number <> 53 And Err.Number <> 76 Then
        ShowError "LoadPayeeMap"
    Else
        LoadPayeeMap = True
    End If
    Exit Function
End Function

Function SavePayeeMap(sFile As String) As Boolean
    Dim vFields As Variant
    Dim iFile As Integer
    Dim oPayee As PayeeMapItem
    Dim sLine As String
    Dim sMsg As String
    
    SavePayeeMap = False
    On Error GoTo nofile
    iFile = FreeFile
    Open sFile For Output Access Write As iFile
    On Error GoTo 0
    vFields = Array("", "", "", "", "", "")
    For Each oPayee In mCol
        vFields(0) = oPayee.Pattern
        vFields(1) = oPayee.Replacement
        vFields(2) = oPayee.Category
        vFields(3) = oPayee.MatchField
        vFields(4) = CStr(oPayee.SIC)
        vFields(5) = IIf(oPayee.NoReturn, "1", "0")
        sLine = MakeLineDelimited(vFields, Delimiter)
        Print #iFile, sLine
    Next
    Close #iFile
    SavePayeeMap = True
    Exit Function
nofile:
' 53: file not found
' 76: path not found
    If Err.Number <> 53 And Err.Number <> 76 Then
'        ShowError "SavePayeeMap"
        sMsg = "Error " & Err.Number & ": " & Err.Description & _
        " on file " & sFile
        MyMsgBox sMsg, vbOKOnly + vbCritical, "SavePayeeMap"
        LogMessage False, True, sMsg
    Else
        SavePayeeMap = True
    End If
    Exit Function
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Clear
' Description:       clear out entire collection
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       19/07/2004-23:29:57
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub Clear()
    Set mCol = New Collection
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       MapSearch
' Description:       Search our map and return the first match!
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/07/2004-23:13:45
'
' Parameters :       sPayee (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function MapSearch(ByVal sMemo As String, ByVal sPayee As String, ByRef sReplacement As String) As PayeeMapItem
' 20050214 CS: added sMemo to allow match to be independent of payee
    Dim re As New RegExp
    Dim Matches As MatchCollection
    Dim m As Match
    Dim i As Long
    Dim sTmp As String
    Dim oItem As PayeeMapItem
    
    On Error GoTo baleout
    re.IgnoreCase = IgnoreCase
    
    For Each oItem In mCol
        re.Pattern = oItem.Pattern
        Select Case oItem.MatchField
        Case mfPayee
            sTmp = sPayee
        Case mfMemo
            sTmp = sMemo
        Case Else
            sTmp = sPayee
            Debug.Assert False
        End Select
        Set Matches = re.Execute(sTmp)
        If Matches.Count > 0 Then
            Set m = Matches(0)
            If Len(oItem.Replacement) > 0 Then
                sTmp = oItem.Replacement
                For i = 1 To m.SubMatches.Count
                    sTmp = Replace(sTmp, "%" & CStr(i), m.SubMatches(i - 1))
                Next
                sReplacement = sTmp
' 20050214 CS: bugfix: empty payee should leave payee alone! added Else branch below
            Else
                sReplacement = sPayee
            End If
            Set MapSearch = oItem
' 20060330 CS: implement NoReturn option. If there is no subsequent match, the routine will "fail" even
' though we have had an initial match.
            If oItem.NoReturn Then
                sPayee = sReplacement
            Else
                Exit Function
            End If
        End If
    Next
goback:
    Set MapSearch = Nothing
    Exit Function
baleout:
    ShowError "MapSearch"
    Resume goback
End Function
