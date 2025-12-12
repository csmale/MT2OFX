VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoRegisterInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : AutoRegisterInfo
'    Project    : MT2OFX
'
'    Description: Information for Auto Registration
'
'    Modified   : $Author: Colin $ $Date: 7/12/06 14:53 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/AutoRegisterInfo.cls 6     7/12/06 14:53 Colin $"
Rem $History: AutoRegisterInfo.cls $
' 
' *****************  Version 6  *****************
' User: Colin        Date: 7/12/06    Time: 14:53
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 3  *****************
' User: Colin        Date: 6/03/05    Time: 0:35
' Updated in $/MT2OFX
'</CSCC>

Public MT940 As Boolean
Public BankID As String
Public BankName As String
Public FormatName As String
Public RecoString As String
Public RecoLines As Integer
Public Structured86 As Boolean
Public SkipEmptyMemoFields As Boolean
Public Extension As String

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Class_Initialize
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       17-Dec-2004-10:42:51
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Class_Initialize()
    Structured86 = False
    RecoLines = 0
    MT940 = False
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       AutoRegisterText
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       17-Dec-2004-11:04:25
'
' Parameters :       sScript (String)
'--------------------------------------------------------------------------------
'</CSCM>
Private Function AutoRegisterText(sScript As String) As Boolean
    AutoRegisterText = False
' insert script at top of list in [text] in ini file

' file extension!
    If Len(Extension) > 0 Then
        If UpdateExtList(Cfg.AutoText, Extension) Then
' rewrite cfg.autotext to inifile
        End If
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       AutoRegisterMT940
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       17-Dec-2004-11:04:46
'
' Parameters :       sScript (String)
'--------------------------------------------------------------------------------
'</CSCM>
Private Function AutoRegisterMT940(sScript As String) As Boolean
    Dim oBankCfg As New BankConfig
    AutoRegisterMT940 = False
    If oBankCfg.Load(BankID) Then
        MsgBox "Bank " & BankID & " is already registered. Please remove the old entry first."
        Exit Function
    End If
    With oBankCfg
        .BankKey = BankID
        .ScriptFile = sScript
        .IDString = BankName
        .Structured86 = Structured86
        .SkipEmptyMemoFields = SkipEmptyMemoFields
    End With
    AutoRegisterMT940 = SaveBankSettingsEx(oBankCfg, BankID)
    If Not AutoRegisterMT940 Then
        Exit Function
    End If
' recognition string!!! bankrules

' file extension!
    If Len(Extension) > 0 Then
        If UpdateExtList(Cfg.AutoMT940, Extension) Then
' rewrite cfg.autotext to inifile
        End If
    End If
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       AutoRegister
' Description:       perform self-registration
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       17-Dec-2004-10:56:18
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function AutoRegister(sScript As String) As Boolean
    Dim bRet As Boolean
    bRet = True
    If Len(BankID) = 0 Then bRet = False
    If Len(BankName) = 0 Then bRet = False
    If MT940 Then
        If RecoLines = 0 Then bRet = False
        If Len(RecoString) = 0 Then bRet = False
    End If
    If Not bRet Then
        AutoRegister = False
        Exit Function
    End If
' remove app.path from front of sscript
    Dim sTmp As String
    If Left$(sScript, Len(App.Path) + 1) = App.Path & "\" Then
        sTmp = Mid$(sScript, Len(App.Path) + 2)
    Else
        sTmp = sScript
    End If
    If MT940 Then
        AutoRegister = AutoRegisterMT940(sTmp)
    Else
        AutoRegister = AutoRegisterText(sTmp)
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       UpdateExtList
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       18-Dec-2004-00:49:06
'
' Parameters :       sExtList (String)
'                    sNewExts (String)
'--------------------------------------------------------------------------------
'</CSCM>
Private Function UpdateExtList(sBase As String, sNewExts As String) As Boolean
    Dim vExts
    Dim sExt As String
    Dim bRet As Boolean
    Dim i As Integer
    
    bRet = False
    vExts = Split(Extension, ",")
    If TypeName(vExts) = "String()" Then
        bRet = False
        For i = LBound(vExts) To UBound(vExts)
            sExt = vExts(i)
            bRet = bRet Or AddToExtensionList(sBase, sExt)
        Next
    End If

    UpdateExtList = bRet
End Function
