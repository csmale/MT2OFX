VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInputExt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input File Extension Mapping"
   ClientHeight    =   4395
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7305
   ControlBox      =   0   'False
   HelpContextID   =   1500
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExtDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   1080
      WhatsThisHelpID =   1503
      Width           =   975
   End
   Begin VB.CommandButton cmdExtNew 
      Caption         =   "New"
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   600
      WhatsThisHelpID =   1502
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cdScript 
      Left            =   6360
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Extension Mapping"
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      WhatsThisHelpID =   1507
      Width           =   7095
      Begin VB.CommandButton cmdSettings 
         Caption         =   "Settings..."
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   960
         WhatsThisHelpID =   1513
         Width           =   1095
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About..."
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   960
         WhatsThisHelpID =   1512
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   4200
         TabIndex        =   5
         Top             =   960
         WhatsThisHelpID =   1509
         Width           =   1095
      End
      Begin VB.CommandButton cmdBrowse 
         Height          =   375
         Left            =   5400
         Picture         =   "frmInputExt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   5880
         TabIndex        =   3
         Top             =   720
         WhatsThisHelpID =   1505
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   5880
         TabIndex        =   2
         Top             =   240
         WhatsThisHelpID =   1504
         Width           =   1095
      End
      Begin VB.TextBox txtScript 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label Label3 
         Caption         =   "Processing Script File:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         WhatsThisHelpID =   1508
         Width           =   2655
      End
   End
   Begin MSComctlLib.ListView lvExts 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   600
      WhatsThisHelpID =   1511
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Extn"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Script"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   3960
      WhatsThisHelpID =   1506
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "File extensions which link non-MT940 files directly to an input script:"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   120
      WhatsThisHelpID =   1501
      Width           =   6375
   End
End
Attribute VB_Name = "frmInputExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : frmInputExt
'    Project    : MT2OFX
'
'    Description: Input Extensions Form
'
'    Modified   : $Author: Colin $ $Date: 6/03/05 0:35 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/frmInputExt.frm 6     6/03/05 0:35 Colin $"
' $History: frmInputExt.frm $
' 
' *****************  Version 6  *****************
' User: Colin        Date: 6/03/05    Time: 0:35
' Updated in $/MT2OFX
'</CSCC>

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       StartEdit
' Description:       start editing the extension mapping
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-18:59:37
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub StartEdit()
    If Len(txtScript) > 0 Then
        Me.cmdReset.Enabled = True
        Me.cmdUpdate.Enabled = True
    Else
        Me.cmdReset.Enabled = False
        Me.cmdUpdate.Enabled = False
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CheckSelection
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-19:00:21
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub CheckSelection()
    If lvExts.SelectedItem Is Nothing Then
        Me.txtScript = ""
        Me.cmdEdit.Enabled = False
        Me.cmdAbout.Enabled = False
        Me.cmdSettings.Enabled = False
        Me.cmdBrowse.Enabled = False
        Me.cmdExtDelete.Enabled = False
    Else
        Me.txtScript = lvExts.SelectedItem.SubItems(1)
        Me.cmdEdit.Enabled = (Len(txtScript) > 0)
        Me.cmdAbout.Enabled = (Len(txtScript) > 0)
        Me.cmdSettings.Enabled = (Len(txtScript) > 0)
        Me.cmdBrowse.Enabled = True
        Me.cmdExtDelete.Enabled = True
    End If
    Me.cmdReset.Enabled = False
    Me.cmdUpdate.Enabled = False
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdAbout_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12-Jan-2005-16:50:23
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

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdBrowse_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-20:28:44
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdBrowse_Click()
    Dim sTmp As String
    
    sTmp = ChooseScript(cdScript, Me.txtScript)
    If sTmp <> "" Then
        Me.txtScript = sTmp
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdClose_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-17:22:04
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
' Procedure  :       cmdEdit_Click
' Description:       Edit a script
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-21:13:28
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
' Procedure  :       cmdExtDelete_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-20:46:37
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdExtDelete_Click()
    Dim sTmp As String
    Dim iTmp As Long
    Dim i As Long
    sTmp = LoadResStringL(130)
    If MyMsgBox(sTmp, vbYesNo + vbDefaultButton2 + vbQuestion, App.Title) = vbYes Then
        DeleteMyString IniSectionTextExtension, Me.lvExts.SelectedItem.Text
        With lvExts
            iTmp = .SelectedItem.Index
            .ListItems.Remove iTmp
            If iTmp > .ListItems.Count Then
                iTmp = .ListItems.Count
            End If
            If iTmp > 0 Then
                Set .SelectedItem = .ListItems(iTmp)
            Else
                Set .SelectedItem = Nothing
            End If
        End With
    End If
    CheckSelection
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdExtNew_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-20:46:28
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdExtNew_Click()
    Dim sTmp As String
    frmNewExt.sExt = ""
    frmNewExt.Show vbModal, Me
    sTmp = UCase$(frmNewExt.sExt)
    If sTmp <> "" Then
        With Me.lvExts.ListItems.Add
            .Text = sTmp
            .SubItems(1) = ""
            .Selected = True
            .EnsureVisible
            DoEvents
            Set Me.lvExts.SelectedItem = Me.lvExts.ListItems(.Index)
            DoEvents
        End With
        CheckSelection
        If Me.cmdBrowse.Enabled Then
            Call cmdBrowse_Click
        End If
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdReset_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-20:26:05
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
' Procedure  :       cmdSettings_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12-Jan-2005-16:49:32
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdSettings_Click()
    Dim sMod As String
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
' Date-Time  :       02/01/2004-20:26:25
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdUpdate_Click()
    Dim sScript As String
    sScript = Me.txtScript
    With Me.lvExts.SelectedItem
        .SubItems(1) = sScript
    End With
    PutMyString IniSectionTextExtension, Me.lvExts.SelectedItem.Text, sScript
    CheckSelection
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        ShowHelpTopic HH_Text_Input_File_Extensions_Window
        KeyCode = 0
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Form_Load
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-17:22:35
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()
    Dim vExts As Variant
    
    LocaliseForm Me
    
    vExts = GetIniKeyList(IniSectionTextExtension)
    Me.lvExts.ListItems.Clear
    Dim sTmp As String
    Dim i As Integer
    For i = LBound(vExts) To UBound(vExts)
        With Me.lvExts.ListItems.Add
            sTmp = UCase$(vExts(i))
            .Text = sTmp
            .SubItems(1) = GetMyString(IniSectionTextExtension, sTmp, "")
            .Selected = False
        End With
    Next
    Set lvExts.SelectedItem = Nothing
    CheckSelection
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       lvExts_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-18:55:10
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub lvExts_Click()
    CheckSelection
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtExtension_Change
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-18:58:54
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtExtension_Change()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtScript_Change
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-18:59:08
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtScript_Change()
    StartEdit
End Sub
