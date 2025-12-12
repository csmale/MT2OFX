VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMatch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MT2OFX MT940 Bank Matching"
   ClientHeight    =   4575
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7440
   ControlBox      =   0   'False
   HelpContextID   =   1600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Match Rule"
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      WhatsThisHelpID =   1606
      Width           =   7215
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   6000
         TabIndex        =   6
         Top             =   720
         WhatsThisHelpID =   1611
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   6000
         TabIndex        =   5
         Top             =   240
         WhatsThisHelpID =   1610
         Width           =   1095
      End
      Begin VB.ComboBox cbBanks 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox txtLines 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtPattern 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   4575
      End
      Begin MSComCtl2.UpDown udLines 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   600
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtLines"
         BuddyDispid     =   196613
         OrigLeft        =   1560
         OrigTop         =   3120
         OrigRight       =   1800
         OrigBottom      =   3495
         Max             =   20
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         Caption         =   "Bank Configuration:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         WhatsThisHelpID =   1609
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Lines"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   600
         WhatsThisHelpID =   1614
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Within:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         WhatsThisHelpID =   1608
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Match pattern:"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         WhatsThisHelpID =   1607
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdDelRule 
      Caption         =   "Remove"
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   2040
      WhatsThisHelpID =   1605
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddRule 
      Caption         =   "New"
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   1560
      WhatsThisHelpID =   1604
      Width           =   1095
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Move Down"
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      Top             =   600
      WhatsThisHelpID =   1603
      Width           =   1095
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Move Up"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   120
      WhatsThisHelpID =   1602
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvMatches 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      WhatsThisHelpID =   1601
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Lines"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Pattern"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Bank"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   4080
      WhatsThisHelpID =   1612
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   4080
      WhatsThisHelpID =   1613
      Width           =   1215
   End
End
Attribute VB_Name = "frmMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : frmMatch
'    Project    : MT2OFX
'
'    Description: MT940 Bank Matching Form
'
'    Modified   : $Author: Colin $ $Date: 2/11/05 23:03 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/frmMatch.frm 8     2/11/05 23:03 Colin $"
' $History: frmMatch.frm $
' 
' *****************  Version 8  *****************
' User: Colin        Date: 2/11/05    Time: 23:03
' Updated in $/MT2OFX
' V3.4 beta 1
'
' *****************  Version 7  *****************
' User: Colin        Date: 18/03/05   Time: 21:57
' Updated in $/MT2OFX
'
' *****************  Version 6  *****************
' User: Colin        Date: 6/03/05    Time: 0:35
' Updated in $/MT2OFX
'</CSCC>

Private iOriginalCount As Long
Const csLink As String = " - "

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       StartEdit
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       17/12/2003-21:51:10
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub StartEdit()
    Me.cmdReset.Enabled = True
    Me.cmdUpdate.Enabled = True
    Me.cmdUp.Enabled = False
    Me.cmdDown.Enabled = False
    Me.cmdAddRule.Enabled = False
    Me.cmdDelRule.Enabled = False
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CheckSelection
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       17/12/2003-21:04:27
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub CheckSelection()
    Dim i As Long
    With lvMatches
        If .SelectedItem Is Nothing Then
            Me.txtLines = ""
            Me.txtLines.Enabled = False
            Me.udLines.Enabled = False
            Me.txtPattern = ""
            Me.txtPattern.Enabled = False
            Me.cbBanks.Enabled = False
            Me.cmdDelRule.Enabled = False
            Me.cmdDown.Enabled = False
            Me.cmdUp.Enabled = False
        Else
            Me.txtLines.Enabled = True
            Me.txtLines = .SelectedItem.SubItems(1)
            Me.txtPattern.Enabled = True
            Me.txtPattern = .SelectedItem.SubItems(2)
            Me.cbBanks.Enabled = True
            If .SelectedItem.SubItems(3) = "" Then
                Me.cbBanks.ListIndex = -1
            Else
                i = FindInCombo(Me.cbBanks, .SelectedItem.SubItems(3))
                If i >= 0 Then
                    Me.cbBanks.ListIndex = i
                Else
                    MyMsgBox LoadResStringL(126)
                    Me.cbBanks.ListIndex = -1
                End If
            End If
            Me.cmdDelRule.Enabled = True
            Me.cmdDown.Enabled = (.SelectedItem.Index < .ListItems.Count)
            Me.cmdUp.Enabled = (.SelectedItem.Index > 1)
        End If
    End With
    Me.cmdReset.Enabled = False
    Me.cmdUpdate.Enabled = False
    Me.cmdAddRule.Enabled = True
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cbBanks_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       17/12/2003-21:51:51
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cbBanks_Click()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdAddRule_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       17/12/2003-23:05:38
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdAddRule_Click()
    With lvMatches.ListItems.Add(, , CStr(lvMatches.ListItems.Count + 1))
        .SubItems(1) = "1"  ' lines
        .SubItems(2) = ""  ' pattern
        .SubItems(3) = ""  ' bank
        Set lvMatches.SelectedItem = lvMatches.ListItems(.Index)
        .EnsureVisible
    End With
    CheckSelection
    txtPattern.SetFocus
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdCancel_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       14/12/2003-20:51:51
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
' Procedure  :       cmdDelRule_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       17/12/2003-22:55:18
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdDelRule_Click()
    Dim sTmp As String
    Dim iTmp As Long
    Dim i As Long
    sTmp = LoadResStringL(127)
    If MyMsgBox(sTmp, vbYesNo + vbDefaultButton2 + vbQuestion, App.Title) = vbYes Then
        With lvMatches
            iTmp = .SelectedItem.Index
            .ListItems.Remove iTmp
            If iTmp > .ListItems.Count Then
                iTmp = .ListItems.Count
            Else
                For i = iTmp To .ListItems.Count
                    .ListItems(i).Text = CStr(i)
                Next
            End If
            If iTmp > 0 Then
                Set .SelectedItem = .ListItems(iTmp)
            Else
                Set .SelectedItem = Nothing
            End If
            cmdOK.Enabled = True
        End With
    End If
    CheckSelection
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdDown_Click
' Description:       move selected item down in list
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       17/12/2003-23:21:11
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdDown_Click()
    Dim iTmp As Long
    Dim i As Long
    Dim sBank As String, sPattern As String, sLines As String
    With lvMatches
        With .SelectedItem
            iTmp = .Index
            sLines = .SubItems(1)
            sPattern = .SubItems(2)
            sBank = .SubItems(3)
        End With
        With .ListItems
            .Remove iTmp  ' remove selected item
                ' now insert before previous item
            With .Add(iTmp + 1, , CStr(iTmp + 1))
                .SubItems(1) = sLines
                .SubItems(2) = sPattern
                .SubItems(3) = sBank
            End With
        End With
        For i = iTmp To .ListItems.Count
            .ListItems(i).Text = CStr(i)
        Next
        Set .SelectedItem = .ListItems(iTmp + 1)
        .SelectedItem.EnsureVisible
    End With
    cmdOK.Enabled = True
    CheckSelection
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdOK_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       14/12/2003-20:52:46
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdOK_Click()
    Dim i As Long
    Dim sTmp As String
    With Me.lvMatches.ListItems
        For i = 1 To .Count
            sTmp = .Item(i).SubItems(3) _
                & "," & .Item(i).SubItems(1) _
                & "," & .Item(i).SubItems(2)
            PutMyString IniSectionBankRules, IniBankRulePrefix & CStr(i), _
                sTmp
        Next
        For i = .Count + 1 To iOriginalCount
            DeleteMyString IniSectionBankRules, IniBankRulePrefix & CStr(i)
        Next
    End With
    Unload Me
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdReset_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       17/12/2003-22:02:55
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
' Procedure  :       cmdUp_Click
' Description:       move selected item up in list
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       17/12/2003-23:21:28
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdUp_Click()
    Dim iTmp As Long
    Dim i As Long
    Dim sBank As String, sPattern As String, sLines As String
    With lvMatches
        With .SelectedItem
            iTmp = .Index
            sLines = .SubItems(1)
            sPattern = .SubItems(2)
            sBank = .SubItems(3)
        End With
        With .ListItems
            .Remove iTmp  ' remove selected item
                ' now insert before previous item
            With .Add(iTmp - 1, , CStr(iTmp - 1))
                .SubItems(1) = sLines
                .SubItems(2) = sPattern
                .SubItems(3) = sBank
            End With
        End With
        For i = iTmp To .ListItems.Count
            .ListItems(i).Text = CStr(i)
        Next
        Set .SelectedItem = .ListItems(iTmp - 1)
        .SelectedItem.EnsureVisible
    End With
    cmdOK.Enabled = True
    CheckSelection
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdUpdate_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       17/12/2003-22:03:54
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdUpdate_Click()
    Dim i As Integer
    Dim sTmp As String
    sTmp = cbBanks
    i = InStr(sTmp, csLink)
    sTmp = Left$(sTmp, i - 1)
    With lvMatches.SelectedItem
        .SubItems(1) = txtLines
        .SubItems(2) = txtPattern
        .SubItems(3) = sTmp
    End With
    cmdOK.Enabled = True
    CheckSelection
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        ShowHelpTopic HH_MT940_Matching_Window
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
' Date-Time  :       14/12/2003-20:53:27
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()
    Dim iBank As Integer
    Dim vArr As Variant
    Dim sTmp As String
    Dim i As Integer
    
    LocaliseForm Me
    cmdOK.Enabled = False
    Me.lvMatches.ListItems.Clear
    iBank = 1
nextbank:
    sTmp = GetMyString(IniSectionBankRules, _
        IniBankRulePrefix & CStr(iBank), "")
    If sTmp = "" Then GoTo donebanks
    vArr = Split(sTmp, ",")
    If Not IsArray(vArr) Then GoTo gonext
    If UBound(vArr) <> 2 Then GoTo gonext
    With Me.lvMatches.ListItems.Add(iBank, , CStr(iBank))
        .SubItems(1) = vArr(1)  ' lines
        .SubItems(2) = vArr(2)  ' pattern
        .SubItems(3) = vArr(0)  ' bank
    End With
gonext:
    iBank = iBank + 1
    GoTo nextbank
donebanks:
    iOriginalCount = iBank - 1
    Dim vTmp As Variant
    vTmp = GetBankSections()
    With cbBanks
        For i = LBound(vTmp) To UBound(vTmp)
            sTmp = vTmp(i)
            sTmp = sTmp & csLink & GetMyString(sTmp, IniBankName, "")
            .AddItem sTmp
        Next
    End With
    CheckSelection
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       lvMatches_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       15/12/2003-01:08:04
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub lvMatches_Click()
    CheckSelection
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       lvMatches_KeyUp
' Description:       check for up/down in listview
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       17/12/2003-21:17:01
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub lvMatches_KeyUp(KeyCode As Integer, Shift As Integer)
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
' Procedure  :       txtLines_Change
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       17/12/2003-21:49:08
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtLines_Change()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtPattern_Change
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       17/12/2003-21:51:28
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtPattern_Change()
    StartEdit
End Sub
