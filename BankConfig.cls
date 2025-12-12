VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BankConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : BankConfig
'    Project    : MT2OFX
'
'    Description: Bank Config (MT940)
'
'    Modified   : $Author: Colin $ $Date: 19/04/08 22:48 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/BankConfig.cls 10    19/04/08 22:48 Colin $"
' $History: BankConfig.cls $
' 
' *****************  Version 10  *****************
' User: Colin        Date: 19/04/08   Time: 22:48
' Updated in $/MT2OFX
' added constructor
'
' *****************  Version 9  *****************
' User: Colin        Date: 7/12/06    Time: 14:53
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 6  *****************
' User: Colin        Date: 6/03/05    Time: 23:41
' Updated in $/MT2OFX
'</CSCC>

Public BankKey As String
Public IDString As String
Public Structured86 As Boolean
Public SkipEmptyMemoFields As Boolean
Public FindPayeeFn As String
Public FindTxnDateFn As String
Public FindServerTimeFn As String
Public IntuitBankID As String
Public ScriptFile As String

Public Changed As Boolean

Const IniStructured86 = "Structured86"
Const IniBankName = "BankName"
Const IniSkipEmptyMemoFields = "SkipEmptyMemoFields"
Const IniFindPayeeFn = "FindPayee"
Const IniFindTxnDateFn = "FindTxnDate"
Const IniFindServerTimeFn = "FindServerTime"
Const IniIntuitBankID = "IntuitBankID"
Const IniScriptFile = "ScriptFile"

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Load
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       17-Dec-2004-16:01:40
'
' Parameters :       sBank (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function Load(sBank As String) As Boolean
    Dim sSection As String
    sSection = sBank
    Load = False
    BankKey = sBank
    IDString = GetMyString(sSection, IniBankName, "")
    If IDString = "" Then
        Structured86 = False
        SkipEmptyMemoFields = False
        FindPayeeFn = ""
        FindTxnDateFn = ""
        FindServerTimeFn = ""
        IntuitBankID = ""
        ScriptFile = ""
    Else
        Structured86 = (GetMyString(sSection, IniStructured86, "0") = "1")
        SkipEmptyMemoFields = (GetMyString(sSection, IniSkipEmptyMemoFields, "0") = "1")
        FindPayeeFn = GetMyString(sSection, IniFindPayeeFn, "")
        FindTxnDateFn = GetMyString(sSection, IniFindTxnDateFn, "")
        FindServerTimeFn = GetMyString(sSection, IniFindServerTimeFn, "")
        IntuitBankID = GetMyString(sSection, IniIntuitBankID, "")
        ScriptFile = GetMyString(sSection, IniScriptFile, "")
        Load = True
    End If
    Changed = False
End Function
' 20041217 CS Remove blank function names from ini-file (depracated anyway)
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       SaveBankSettingsEx
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       17-Dec-2004-16:03:17
'
' Parameters :       bc (BankConfig)
'                    sSection (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function Save(sSection As String) As Boolean
    PutMyString sSection, IniBankName, IDString
    PutMyString sSection, IniStructured86, IIf(Structured86, "1", "0")
    PutMyString sSection, IniSkipEmptyMemoFields, IIf(SkipEmptyMemoFields, "1", "0")
    If Len(FindPayeeFn) > 0 Then
        PutMyString sSection, IniFindPayeeFn, FindPayeeFn
    Else
        DeleteMyString sSection, IniFindPayeeFn
    End If
    If Len(FindTxnDateFn) > 0 Then
        PutMyString sSection, IniFindTxnDateFn, FindTxnDateFn
    Else
        DeleteMyString sSection, IniFindTxnDateFn
    End If
    If Len(FindServerTimeFn) > 0 Then
        PutMyString sSection, IniFindServerTimeFn, FindServerTimeFn
    Else
        DeleteMyString sSection, IniFindServerTimeFn
    End If
    PutMyString sSection, IniScriptFile, ScriptFile
    Changed = False
End Function

Private Sub Class_Initialize()
    BankKey = ""
    IDString = ""
    Structured86 = False
    SkipEmptyMemoFields = False
    FindPayeeFn = ""
    FindTxnDateFn = ""
    FindServerTimeFn = ""
    IntuitBankID = ""
    ScriptFile = ""
    Changed = False
End Sub
