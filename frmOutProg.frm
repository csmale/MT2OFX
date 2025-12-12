VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOutProg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Output Program"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optProgNone 
      Caption         =   "Do nothing"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   600
      WhatsThisHelpID =   1352
      Width           =   5295
   End
   Begin VB.ComboBox cbFileTypes 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog cdChooseProgram 
      Left            =   120
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   375
      Left            =   5760
      Picture         =   "frmOutProg.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtProgPath 
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1560
      Width           =   4815
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   2040
      WhatsThisHelpID =   1356
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   2040
      WhatsThisHelpID =   1355
      Width           =   1095
   End
   Begin VB.OptionButton optProgUser 
      Caption         =   "Open it with the following program:"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      WhatsThisHelpID =   1354
      Width           =   5295
   End
   Begin VB.OptionButton optProgDefault 
      Caption         =   "Open it with program configured in Windows"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      WhatsThisHelpID =   1353
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Output File Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      WhatsThisHelpID =   1351
      Width           =   1695
   End
End
Attribute VB_Name = "frmOutProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : frmOutProg
'    Project    : MT2OFX
'
'    Description:
'
'    Modified   : $Author: Colin $ $Date: 18/03/05 21:57 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/frmOutProg.frm 7     18/03/05 21:57 Colin $"
' $History: frmOutProg.frm $
' 
' *****************  Version 7  *****************
' User: Colin        Date: 18/03/05   Time: 21:57
' Updated in $/MT2OFX
'
' *****************  Version 6  *****************
' User: Colin        Date: 6/03/05    Time: 23:41
' Updated in $/MT2OFX
'</CSCC>

Public OptCfg As ProgramConfig
Public Cancelled As Boolean
Private sFileType As String

'// Windows Registry Messages
Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

'// Windows Error Messages
Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

'// Windows Security Messages
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const KEY_ALL_ACCESS = &H3F
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const REG_OPTION_NON_VOLATILE = 0

'// Windows Registry API calls
Private Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal hKey As Long) As Long

Private Declare Function RegOpenKeyEx _
 Lib "advapi32.dll" Alias "RegOpenKeyExA" _
(ByVal hKey As Long, _
 ByVal lpSubKey As String, _
 ByVal ulOptions As Long, _
 ByVal samDesired As Long, _
 phkResult As Long) As Long

Private Declare Function RegQueryValueEx _
 Lib "advapi32.dll" Alias "RegQueryValueExA" _
 (ByVal hKey As Long, _
  ByVal lpValueName As String, _
  ByVal lpReserved As Long, _
  lpType As Long, _
  lpData As Any, _
  lpcbData As Long) As Long
  

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cbFileTypes_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       11/02/2005-16:58:05
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cbFileTypes_Click()
    Dim sTmp As String
    Dim iComma As Integer
    Dim vArr As Variant
    Dim iOpt As Integer

    If sFileType <> "" Then
    
    If optProgNone Then
        sTmp = "0"
    ElseIf optProgDefault Then
        sTmp = "1"
    ElseIf optProgUser Then
        sTmp = "2," & txtProgPath
    Else
        sTmp = "1"
    End If
    
    Select Case sFileType
    Case "OFX"
        OptCfg.OFXExportTo = sTmp
    Case "OFC"
        OptCfg.OFCExportTo = sTmp
    Case "QIF"
        OptCfg.QIFExportTo = sTmp
    Case "QFX"
        OptCfg.QFXExportTo = sTmp
    End Select
    End If

    sFileType = Left$(cbFileTypes, 3)
    
    sTmp = GetFileHandler(sFileType)
    If sTmp = "" Then
        iOpt = 1
    Else
        vArr = Split(sTmp, ",")
        If IsNumeric(vArr(0)) Then
            iOpt = CInt(vArr(0))
            If iOpt < 0 Or iOpt > 2 Then
                iOpt = 1
            End If
        Else
            iOpt = 1
        End If
        If iOpt = 2 Then    ' user defined prog
            If UBound(vArr) = 0 Then
                iOpt = 1
            End If
        End If
    End If
    Select Case iOpt
    Case 0
        optProgNone = True
        txtProgPath = ""
    Case 1
        optProgDefault = True
        txtProgPath = ""
    Case 2
        txtProgPath = vArr(1)
        optProgUser = True
    End Select
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdBrowse_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       11/02/2005-16:40:48
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdBrowse_Click()
    With cdChooseProgram
        .DialogTitle = "Choose Program"
        .Filter = "Executable Programs|*.exe|All Programs|*.exe;*.cmd;*.bat|All Files|*.*"
        .FilterIndex = 1
        .Flags = cdlOFNExplorer + cdlOFNFileMustExist + _
            cdlOFNHideReadOnly + cdlOFNLongNames
        If txtProgPath = "" Then
            .InitDir = GetSpecialFolder(CSIDL_PROGRAM_FILES)
            .FileName = ""
        Else
            .InitDir = txtProgPath
            .FileName = txtProgPath
        End If
        .ShowOpen
        If .FileName <> "" Then
            txtProgPath = .FileName
            optProgUser = True
        Else
            optProgDefault = True
        End If
    End With
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdCancel_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       28-Jan-2005-22:44:51
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdCancel_Click()
    Unload Me
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdOK_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       11/02/2005-21:57:56
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdOK_Click()
    Cancelled = False
    OptCfg.Changed = True
    Call cbFileTypes_Click  ' get last changes
    Unload Me
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Form_Load
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       11/02/2005-16:26:43
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()
    Dim aExt As Variant
    If OptCfg Is Nothing Then
        Debug.Print "frmOutProg: no options"
        Debug.Assert False
    End If
    LocaliseForm Me, 1350
    With cbFileTypes
        .Clear
        For Each aExt In Array("OFX", "OFC", "QIF", "QFX")
            .AddItem aExt & " (" & FileTypeString(CStr(aExt)) & ")"
        Next
        sFileType = ""
        .ListIndex = FindInCombo(cbFileTypes, OptCfg.OutputFileType)
    End With
    Cancelled = True
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       FileTypeString
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       11/02/2005-16:31:02
'
' Parameters :       sExt (String)
'--------------------------------------------------------------------------------
'</CSCM>
Private Function FileTypeString(sExt As String) As String
    Dim sProgID As String
    sProgID = GetRegString(HKEY_CLASSES_ROOT, "." & sExt, "")
    If sProgID = "" Then
        FileTypeString = ""
    Else
        FileTypeString = GetRegString(HKEY_CLASSES_ROOT, sProgID, "")
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetFileHandler
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       11/02/2005-17:02:06
'
' Parameters :       sExt (String)
'--------------------------------------------------------------------------------
'</CSCM>
Private Function GetFileHandler(sExt As String) As String
    Select Case sExt
    Case "OFX"
        GetFileHandler = OptCfg.OFXExportTo
    Case "OFC"
        GetFileHandler = OptCfg.OFCExportTo
    Case "QIF"
        GetFileHandler = OptCfg.QIFExportTo
    Case "QFX"
        GetFileHandler = OptCfg.QFXExportTo
    Case Else
        Debug.Print "GetFileHandler: Unexpected extension: " & sExt
        Debug.Assert False
        GetFileHandler = ""
    End Select
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       optProgUser_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       11/02/2005-21:46:02
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub optProgUser_Click()
    If optProgUser Then
        If txtProgPath = "" Then
            Call cmdBrowse_Click
        End If
    End If
End Sub
