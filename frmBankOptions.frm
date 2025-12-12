VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBankOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MT2OFX MT940 Bank Options"
   ClientHeight    =   4770
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6480
   ControlBox      =   0   'False
   HelpContextID   =   1400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmBank 
      Caption         =   "Bank Properties"
      Height          =   2175
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      WhatsThisHelpID =   1412
      Width           =   6255
      Begin VB.CommandButton cmdSettings 
         Caption         =   "Settings..."
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   1200
         WhatsThisHelpID =   1419
         Width           =   1335
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About..."
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         WhatsThisHelpID =   1418
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   3480
         TabIndex        =   7
         Top             =   1200
         WhatsThisHelpID =   1410
         Width           =   975
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   5040
         TabIndex        =   12
         Top             =   1080
         WhatsThisHelpID =   1405
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   5040
         TabIndex        =   11
         Top             =   600
         WhatsThisHelpID =   1404
         Width           =   1095
      End
      Begin VB.TextBox txtBankName 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   3255
      End
      Begin VB.CheckBox chkSkipEmptyFields 
         Caption         =   "Ignore empty description fields"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         WhatsThisHelpID =   1409
         Width           =   3855
      End
      Begin VB.CheckBox chkStructured86 
         Caption         =   "Use structured :86: records"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         WhatsThisHelpID =   1408
         Width           =   3855
      End
      Begin VB.CommandButton cmdScriptBrowse 
         Height          =   375
         Left            =   4560
         Picture         =   "frmBankOptions.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtScript 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label2 
         Caption         =   "Bank Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         WhatsThisHelpID =   1406
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Processing script file:"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   600
         WhatsThisHelpID =   1407
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdMatching 
      Caption         =   "Matching..."
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      WhatsThisHelpID =   1411
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog cdScript 
      Left            =   5280
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBankAdd 
      Caption         =   "New..."
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      WhatsThisHelpID =   1402
      Width           =   1095
   End
   Begin VB.CommandButton cmdBankDelete 
      Caption         =   "Remove"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   600
      WhatsThisHelpID =   1403
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   14
      Top             =   4320
      WhatsThisHelpID =   1401
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvBanks 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      WhatsThisHelpID =   1417
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "BIC"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bank Name"
         Object.Width           =   5292
      EndProperty
   End
End
Attribute VB_Name = "frmBankOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : frmBankOptions
'    Project    : MT2OFX
'
'    Description: Bank Options Form
'
'    Modified   : $Author: Colin $ $Date: 7/12/06 14:57 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/frmBankOptions.frm 13    7/12/06 14:57 Colin $"
' $History: frmBankOptions.frm $
' 
' *****************  Version 13  *****************
' User: Colin        Date: 7/12/06    Time: 14:57
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 10  *****************
' User: Colin        Date: 6/05/05    Time: 23:09
' Updated in $/MT2OFX
' Control width changes for German localisation
'
' *****************  Version 9  *****************
' User: Colin        Date: 18/03/05   Time: 21:57
' Updated in $/MT2OFX
'
' *****************  Version 8  *****************
' User: Colin        Date: 6/03/05    Time: 0:35
' Updated in $/MT2OFX
'</CSCC>

Private bc As New BankConfig

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       StartEdit
' Description:       Start editing bank details
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       01/01/2004-23:00:44
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub StartEdit()
    If Not Me.cmdUpdate.Enabled Then
        With Me
            .cmdReset.Enabled = True
            .cmdUpdate.Enabled = True
        End With
    End If
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CheckSelection
' Description:       enable/disable things according to selection
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       10/12/2003-23:29:00
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub CheckSelection()
    If lvBanks.SelectedItem Is Nothing Then
        With Me
            .cmdBankDelete.Enabled = False
            .txtBankName = ""
            .txtBankName.Enabled = False
            .txtScript = ""
            .txtScript.Enabled = False
            .cmdScriptBrowse.Enabled = False
            .chkSkipEmptyFields.Enabled = False
            .chkStructured86.Enabled = False
            .cmdBankDelete.Enabled = False
            .cmdScriptBrowse.Enabled = False
            .cmdUpdate.Enabled = False
            .cmdReset.Enabled = False
            .cmdEdit.Enabled = False
            .cmdAbout.Enabled = False
            .cmdSettings.Enabled = False
        End With
    Else
        If bc.Load(lvBanks.SelectedItem.Text) Then
            With Me
                .cmdBankDelete.Enabled = True
                .txtBankName.Enabled = True
                .txtBankName = bc.IDString
                .txtScript = bc.ScriptFile
                .txtScript.Enabled = True
                .cmdScriptBrowse.Enabled = True
                .chkSkipEmptyFields.Enabled = True
                .chkSkipEmptyFields.Value = IIf(bc.SkipEmptyMemoFields, vbChecked, vbUnchecked)
                .chkStructured86.Enabled = True
                .chkStructured86.Value = IIf(bc.Structured86, vbChecked, vbUnchecked)
                .cmdBankDelete.Enabled = True
                .cmdScriptBrowse.Enabled = True
                .cmdUpdate.Enabled = False
                .cmdReset.Enabled = False
                .cmdEdit.Enabled = (Len(txtScript) > 0)
                .cmdAbout.Enabled = (Len(txtScript) > 0)
                .cmdSettings.Enabled = (Len(txtScript) > 0)
            End With
        End If
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       chkSkipEmptyFields_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       01/01/2004-23:02:22
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub chkSkipEmptyFields_Click()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       chkStructured86_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       01/01/2004-23:00:18
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub chkStructured86_Click()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdAbout_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       14-Jan-2005-22:14:24
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdAbout_Click()
    If txtScript = "" Then
        Exit Sub
    End If
    Call ScriptShowAbout(txtScript)
End Sub

Private Sub cmdBankAdd_Click()
    frmNewBank.Show vbModal, Me
    If Len(frmNewBank.sBankCode) > 0 Then
        With bc
            .BankKey = frmNewBank.sBankCode
            .IDString = frmNewBank.sBankName
            .ScriptFile = ""
            .SkipEmptyMemoFields = False
            .Structured86 = False
        End With
        SaveBankSettingsEx bc, bc.BankKey
        With lvBanks.ListItems.Add
            .Text = bc.BankKey
            .SubItems(1) = bc.IDString
            .Selected = True
            .EnsureVisible
        End With
        CheckSelection
        Call cmdScriptBrowse_Click
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       BankUsedInMatching
' Description:       Check matching rules to see if a bank is a target
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-15:08:35
'
' Parameters :       sBank (String)
'--------------------------------------------------------------------------------
'</CSCM>
Private Function BankUsedInMatching(sBank As String) As Boolean
    Dim iBank As Integer
    Dim sTmp As String
    Dim vArr As Variant
    iBank = 1
nextbank:
    sTmp = GetMyString(IniSectionBankRules, _
        IniBankRulePrefix & CStr(iBank), "")
    If sTmp = "" Then GoTo donebanks
    vArr = Split(sTmp, ",")
    If Not IsArray(vArr) Then GoTo gonext
    If UBound(vArr) <> 2 Then GoTo gonext
    If vArr(0) = sBank Then
        BankUsedInMatching = True
        Exit Function
    End If
gonext:
    iBank = iBank + 1
    GoTo nextbank
donebanks:
    BankUsedInMatching = False
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdBankDelete_Click
' Description:       Handle click on Delete button
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-15:50:18
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdBankDelete_Click()
    Dim sBank As String
    sBank = Me.lvBanks.SelectedItem.Text
    If BankUsedInMatching(sBank) Then
        MyMsgBox LoadResStringL(1415), _
            vbOKOnly + vbCritical, LoadResStringL(1414)
        Exit Sub
    End If
    If MyMsgBox(LoadResStringLEx(1416, sBank), _
        vbYesNo + vbQuestion, LoadResStringL(1414)) = vbYes Then
        DeleteMySection sBank
        Me.lvBanks.ListItems.Remove Me.lvBanks.SelectedItem.Index
        Set Me.lvBanks.SelectedItem = Nothing
        CheckSelection
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdEdit_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-21:16:38
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdEdit_Click()
    EditScript Me.txtScript
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdMatching_Click
' Description:       handle click on Matching...
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       14/12/2003-20:50:32
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdMatching_Click()
    frmMatch.Show vbModal, Me
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdClose_Click
' Description:       handle click on OK button
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       11/12/2003-21:59:59
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdClose_Click()
    Dim iAns As Integer
    If Me.cmdUpdate.Enabled Then
' apparently unsaved changes
        iAns = MyMsgBox(LoadResStringL(132), vbYesNoCancel + vbQuestion, App.Title)
        If iAns = vbYes Then
            Call cmdUpdate_Click
        ElseIf iAns = vbCancel Then
            Exit Sub
        End If
    End If
    Unload Me
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdReset_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       01/01/2004-23:13:45
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdReset_Click()
    CheckSelection
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdScriptBrowse_Click
' Description:       handle click on browse button
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       11/12/2003-22:11:23
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdScriptBrowse_Click()
    Dim iLen As Long
    Dim sTmp As String
    sTmp = ChooseScript(cdScript, Me.txtScript)
    If sTmp <> "" Then
        Me.txtScript = sTmp
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdSettings_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       14-Jan-2005-22:02:07
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdSettings_Click()
    If txtScript = "" Then
        Exit Sub
    End If
    Call ScriptShowSettings(txtScript)
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdUpdate_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       01/01/2004-23:16:47
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdUpdate_Click()
    With bc
        .IDString = Me.txtBankName
        .ScriptFile = Me.txtScript
        .SkipEmptyMemoFields = (Me.chkSkipEmptyFields.Value = vbChecked)
        .Structured86 = (Me.chkStructured86.Value = vbChecked)
    End With
    SaveBankSettingsEx bc, bc.BankKey
    lvBanks.SelectedItem.SubItems(1) = bc.IDString
    CheckSelection
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Form_KeyDown
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/08/2004-21:58:33
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        ShowHelpTopic HH_MT940_Bank_Options_Window
        KeyCode = 0
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       lvBanks_Click
' Description:       Process click on list of banks
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       10/12/2003-23:07:49
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub lvBanks_Click()
    CheckSelection
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       lvBanks_KeyUp
' Description:       handle keypress in bank list
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       10/12/2003-23:33:02
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub lvBanks_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown _
    Or KeyCode = vbKeyUp _
    Or KeyCode = vbKeyPageDown _
    Or KeyCode = vbKeyPageUp Then
        CheckSelection
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Form_Load
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       11/12/2003-21:49:47
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()
    Dim vTmp As Variant
    
    LocaliseForm Me
    
    vTmp = GetBankSections()
    Dim i As Long
    Dim sTmp As String
    lvBanks.ListItems.Clear
' 20061130 CS: stop aborts on empty list
    If UBound(vTmp) >= LBound(vTmp) Then
        With lvBanks
            For i = LBound(vTmp) To UBound(vTmp)
                With .ListItems.Add
                    sTmp = vTmp(i)
                    .Text = sTmp
                    .SubItems(1) = GetMyString(sTmp, "BankName", "")
                    .Selected = False
                End With
            Next
        End With
    End If
    Set lvBanks.SelectedItem = Nothing
    CheckSelection
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtBankName_Change
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       01/01/2004-23:02:51
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtBankName_Change()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtScript_Change
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       01/01/2004-23:03:08
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtScript_Change()
    StartEdit
End Sub
