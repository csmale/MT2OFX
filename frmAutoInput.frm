VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAutoInput 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Text Input Auto-Recognition"
   ClientHeight    =   6600
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7305
   ControlBox      =   0   'False
   HelpContextID   =   1300
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCollect 
      Caption         =   "Collect"
      Height          =   360
      Left            =   6000
      TabIndex        =   22
      Top             =   1920
      WhatsThisHelpID =   1360
      Width           =   1215
   End
   Begin VB.CommandButton cmdWizard 
      Caption         =   "Wizard..."
      Height          =   375
      Left            =   2280
      TabIndex        =   21
      Top             =   6120
      WhatsThisHelpID =   1359
      Width           =   1095
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "Auto!"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdScript 
      Left            =   6720
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Script"
      Height          =   2295
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   7095
      Begin VB.TextBox txtScriptExts 
         Height          =   285
         Left            =   3480
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmdSettings 
         Caption         =   "Settings..."
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   1800
         WhatsThisHelpID =   1315
         Width           =   1095
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About..."
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         WhatsThisHelpID =   1314
         Width           =   975
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   5880
         TabIndex        =   8
         Top             =   720
         WhatsThisHelpID =   1308
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   5880
         TabIndex        =   7
         Top             =   240
         WhatsThisHelpID =   1307
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   4080
         TabIndex        =   11
         Top             =   1800
         WhatsThisHelpID =   1309
         Width           =   1095
      End
      Begin VB.CommandButton cmdBrowse 
         Height          =   375
         Left            =   5280
         Picture         =   "frmAutoInput.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtScript 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   480
         Width           =   5655
      End
      Begin VB.Label Label4 
         Caption         =   "Separate extensions with commas."
         Height          =   495
         Left            =   3480
         TabIndex        =   19
         Top             =   1200
         WhatsThisHelpID =   1358
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Extensions which will automatically select this script without going through the recognition process:"
         Height          =   735
         Left            =   120
         TabIndex        =   18
         Top             =   840
         WhatsThisHelpID =   1357
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Script Name"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         WhatsThisHelpID =   1310
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Remove"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   3240
      WhatsThisHelpID =   1306
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   2760
      WhatsThisHelpID =   1305
      Width           =   1215
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Move Down"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   1440
      WhatsThisHelpID =   1304
      Width           =   1215
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Move Up"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   960
      WhatsThisHelpID =   1303
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvScripts 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   960
      WhatsThisHelpID =   1311
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Script"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Extensions"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   6120
      WhatsThisHelpID =   1301
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   6120
      WhatsThisHelpID =   1302
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   $"frmAutoInput.frx":0342
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   120
      WhatsThisHelpID =   1313
      Width           =   6975
   End
End
Attribute VB_Name = "frmAutoInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : frmAutoInput
'    Project    : MT2OFX
'
'    Description: Auto Input Recognition Form
'
'    Modified   : $Author: Colin $ $Date: 30/01/11 11:33 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/frmAutoInput.frm 22    30/01/11 11:33 Colin $"
' $History: frmAutoInput.frm $
' 
' *****************  Version 22  *****************
' User: Colin        Date: 30/01/11   Time: 11:33
' Updated in $/MT2OFX
' added Collect button
'
' *****************  Version 20  *****************
' User: Colin        Date: 19/04/08   Time: 22:17
' Updated in $/MT2OFX
'
' *****************  Version 19  *****************
' User: Colin        Date: 7/12/06    Time: 14:55
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 16  *****************
' User: Colin        Date: 2/11/05    Time: 23:03
' Updated in $/MT2OFX
' V3.4 beta 1
'
' *****************  Version 15  *****************
' User: Colin        Date: 8/05/05    Time: 12:43
' Updated in $/MT2OFX
' V3.3.8
'
' *****************  Version 14  *****************
' User: Colin        Date: 23/03/05   Time: 22:13
' Updated in $/MT2OFX
' Leaving for Ireland!
'
' *****************  Version 13  *****************
' User: Colin        Date: 18/03/05   Time: 21:57
' Updated in $/MT2OFX
'
' *****************  Version 12  *****************
' User: Colin        Date: 6/03/05    Time: 0:35
' Updated in $/MT2OFX
'</CSCC>

Dim iOriginalCount As Long
Dim gsClose As String
Dim gsCancel As String
Dim bInsertInProgress As Boolean

Private Function MakeLabel(i As Long) As String
    MakeLabel = Right$("   " & CStr(i), 3)
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       StartEdit
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       03/01/2004-01:04:31
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub StartEdit()
    cmdCancel.Caption = gsCancel
' 20050215 CS: kill the new button for now or we can get a list full of empties
    cmdNew.Enabled = False
' Reset can only be used if there's something in the listbox!
    cmdReset.Enabled = True
    cmdUpdate.Enabled = (txtScript <> "")
    cmdWizard.Enabled = False
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CheckSelection
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       03/01/2004-00:46:20
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub CheckSelection()
    Me.lvScripts.Enabled = Not bInsertInProgress
    If lvScripts.SelectedItem Is Nothing Then
        Me.cmdUp.Enabled = False
        Me.cmdDown.Enabled = False
        Me.cmdDelete.Enabled = False
        Me.cmdBrowse.Enabled = False
    Else
        Me.cmdDown.Enabled = (lvScripts.SelectedItem.Index < lvScripts.ListItems.Count)
        Me.cmdUp.Enabled = (lvScripts.SelectedItem.Index > 1)
        Me.cmdDelete.Enabled = True
        Me.txtScript = lvScripts.SelectedItem.SubItems(1)
        Me.txtScriptExts = lvScripts.SelectedItem.SubItems(2)
        Me.cmdEdit.Enabled = (Len(txtScript) > 0)
        Me.cmdBrowse.Enabled = True
    End If
    Me.cmdReset.Enabled = False
    Me.cmdUpdate.Enabled = False
    Me.cmdWizard.Enabled = True
' 20050215 CS: reenable New if we are not editing an entry
    Me.cmdNew.Enabled = True
    Me.cmdAbout.Enabled = (Len(txtScript) > 0)
    Me.cmdSettings.Enabled = (Len(txtScript) > 0)
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdAbout_Click
' Description:       handle click on "About" button
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       17/11/2004-11:00:30
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
' Procedure  :       cmdAuto_Click
' Description:       Automatically search for and add scripts
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       11/2/2005-18:53:22
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdAuto_Click()
    MsgBox "Automatically search for and add scripts"
    Dim xScripts As Scripting.Dictionary
    Dim sFile As String
    Dim sFileLC As String
    Dim i As Long
    Dim iIndex As Long: iIndex = lvScripts.ListItems.Count
    Dim sTmp As String
    Dim vTmp As Variant: vTmp = GetBankSections()

    sFile = Dir(Cfg.ScriptPath & "\*.vbs")
    Do While Len(sFile) > 0
        sFileLC = LCase$(sFile)
' ignore our library
        If sFileLC = "mt2ofx.vbs" Then GoTo nextfile
' check against known non-mt940 scripts
        For i = 1 To Me.lvScripts.ListItems.Count
            If sFileLC = LCase$(lvScripts.ListItems(i).SubItems(1)) Then GoTo nextfile
        Next
' check against known mt940 scripts
        For i = LBound(vTmp) To UBound(vTmp)
            sTmp = vTmp(i)
            If sFileLC = LCase$(GetMyString(sTmp, "ScriptFile", "")) Then GoTo nextfile
        Next
' script really is unknown!
        i = MsgBox("Do you want to add script: " & sFile & "?", vbYesNoCancel, "New Script Found")
        Select Case i
        Case vbYes
        Case vbCancel
            Exit Sub
        Case Else
            GoTo nextfile
        End Select
        i = lvScripts.ListItems.Count + 1
        With lvScripts.ListItems.Add(, , MakeLabel(i))
            .SubItems(1) = sFile  ' script name
            .SubItems(2) = ""  ' extensions
            .Checked = True ' assume active
        End With
nextfile:
        sFile = Dir
    Loop
    Set lvScripts.SelectedItem = lvScripts.ListItems(iIndex + 1)
    lvScripts.SelectedItem.EnsureVisible
    MsgBox "Done."
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdBrowse_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       03/01/2004-01:06:40
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdBrowse_Click()
    Dim iLen As Long
    Dim sTmp As String
    
    sTmp = ChooseScript(cdScript, Me.txtScript)
    If sTmp <> "" Then
        Me.txtScript = sTmp
        StartEdit
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdCancel_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       03/01/2004-00:42:42
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
' Procedure  :       cmdCollect_Click
' Description:       Gather all selected scripts at the top of the list
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       2/9/2010-14:18:56
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdCollect_Click()
    Dim iTmp As Long
    Dim i As Long
    Dim sScript As String
    Dim bChecked As Boolean
    Dim iStart As Long, iFirstUnchecked As Long, iFirstChecked As Long
    
    iStart = 1
    lvScripts.Sorted = False
    Do While True
' find first unchecked item
        iFirstUnchecked = -1
        For i = iStart To lvScripts.ListItems.Count
            If Not lvScripts.ListItems(i).Checked Then
                iFirstUnchecked = i
                Exit For
            End If
        Next
        If iFirstUnchecked < 0 Then
' no unchecked item found
            Exit Do
        End If

' iFirstUnchecked is pointing to the first unchecked item
' scan down to find first checked item from this point
        iFirstChecked = -1
        For i = iFirstUnchecked To lvScripts.ListItems.Count
            If lvScripts.ListItems(i).Checked Then
                iFirstChecked = i
                Exit For
            End If
        Next
        If iFirstChecked < 0 Then
' no more checked items
            Exit Do
        End If

' now we have iFirstUnchecked as the insertion point
' and iFirstChecked needs to go there
        With lvScripts.ListItems(iFirstChecked)
            sScript = .SubItems(1)
            bChecked = .Checked
        End With

' insert an item before the first unchecked and copy first checked to here
        With lvScripts.ListItems
                ' insert before previous item
            With .Add(iFirstUnchecked, , MakeLabel(iFirstUnchecked))
                .SubItems(1) = sScript
                .Checked = bChecked
            End With
' remove first checked from original place (NB: index is one higher as we have just inserted an item)
            .Remove iFirstChecked + 1
        End With
' go back and loop starting from where we got to
        iStart = iFirstUnchecked + 1
    Loop

' reorg of listview complete, tidy up the labels
    With lvScripts
        For i = 1 To .ListItems.Count
            .ListItems(i).Text = MakeLabel(i)
        Next
        Set .SelectedItem = .ListItems(1)
        .SelectedItem.EnsureVisible
    End With
    
    lvScripts.Sorted = True
    cmdOK.Enabled = True
    CheckSelection
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdDelete_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       03/01/2004-01:10:35
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdDelete_Click()
    Dim sTmp As String
    Dim iTmp As Long
    Dim i As Long
    sTmp = LoadResStringL(129)
' 20050322 CS: now looks at the listview to decide whether to ask user
    If lvScripts.SelectedItem.SubItems(1) = "" Then
        i = vbYes
    Else
        i = MyMsgBox(sTmp, vbYesNo + vbDefaultButton2 + vbQuestion, App.Title)
    End If
    If i = vbYes Then
        With lvScripts
            iTmp = .SelectedItem.Index
            .ListItems.Remove iTmp
            If iTmp > .ListItems.Count Then
                iTmp = .ListItems.Count
            Else
                For i = iTmp To .ListItems.Count
                    .ListItems(i).Text = MakeLabel(i)
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
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       03/01/2004-17:20:48
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdDown_Click()
    Dim iTmp As Long
    Dim i As Long
    Dim sScript As String
    Dim bChecked As Boolean
    
    With lvScripts
        With .SelectedItem
            iTmp = .Index
            sScript = .SubItems(1)
            bChecked = .Checked
        End With
        With .ListItems
            .Remove iTmp  ' remove selected item
                ' now insert before previous item
            With .Add(iTmp + 1, , MakeLabel(iTmp + 1))
                .SubItems(1) = sScript
                .Checked = bChecked
            End With
        End With
        For i = iTmp To .ListItems.Count
            .ListItems(i).Text = MakeLabel(i)
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
' Procedure  :       cmdEdit_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       03/01/2004-01:05:45
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
' Procedure  :       cmdNew_Click
' Description:       add a new entry to the list of scripts
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       03/01/2004-01:10:03
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdNew_Click()
    Dim i As Long
    With lvScripts
        With .ListItems.Add(1, , MakeLabel(1))
            .SubItems(1) = ""  ' script name
            .SubItems(2) = ""  ' extensions
            .Checked = True ' assume active
            Set lvScripts.SelectedItem = lvScripts.ListItems(.Index)
            .EnsureVisible
        End With
        For i = 2 To .ListItems.Count
            .ListItems(i).Text = MakeLabel(i)
        Next
    End With
    bInsertInProgress = True
    CheckSelection
    If cmdBrowse.Enabled Then
        Call cmdBrowse_Click
    End If
' 20050322 CS: Always enable Reset button
    cmdReset.Enabled = True
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdOK_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       03/01/2004-00:42:25
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdOK_Click()
    Dim i As Long
    Dim sTmp As String
    Dim vExts As Variant
    Dim sExt As String
    Dim iExt As Integer
    Dim dictExts As New Scripting.Dictionary
    Dim sMsg As String
    
' scan through the list to check there are no duplicate extensions
    With Me.lvScripts.ListItems
        For i = 1 To .Count
            sTmp = .Item(i).SubItems(1)
            If .Item(i).SubItems(2) <> "" Then
                vExts = Split(.Item(i).SubItems(2), ",")
                For iExt = LBound(vExts) To UBound(vExts)
                    sExt = vExts(iExt)
                    If dictExts.Exists(sExt) Then
                        sMsg = LoadResStringLEx(1317, sExt, sTmp, dictExts(sExt))
                        MyMsgBox sMsg, vbOKOnly + vbCritical, "MT2OFX"
                        Exit Sub
                    Else
                        dictExts.Item(sExt) = sTmp
                    End If
                Next
            End If
        Next
    End With
    
' save the script list
    With Me.lvScripts.ListItems
        For i = 1 To .Count
            sTmp = .Item(i).SubItems(1)
            PutMyString IniSectionText, IniTextScriptPrefix & CStr(i), _
                IIf(.Item(i).Checked, "1", "0") & "," & sTmp
        Next
        For i = .Count + 1 To iOriginalCount
            DeleteMyString IniSectionText, IniTextScriptPrefix & CStr(i)
        Next
    End With
' delete all extension maps, leaving section header in place
    ClearMySection IniSectionTextExtension
' add all (new) extension maps
    For Each vExts In dictExts
        PutMyString IniSectionTextExtension, CStr(vExts), dictExts(vExts)
    Next
    Unload Me
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdReset_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       03/01/2004-21:47:36
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdReset_Click()
    Dim bOK As Boolean
    bOK = cmdOK.Enabled
' if the entry is empty, delete it from the list!
    If lvScripts.SelectedItem.SubItems(1) = "" Then
        If cmdDelete.Enabled Then
            Call cmdDelete_Click
            cmdOK.Enabled = bOK
        End If
    End If
    bInsertInProgress = False
    CheckSelection
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdSettings_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       13/11/2004-00:27:36
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

Private Sub cmdWizard_Click()
    frmQuickStart.Show vbModal
    LoadScriptList
    cmdOK.Enabled = True
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdUp_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       03/01/2004-17:16:04
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdUp_Click()
    Dim iTmp As Long
    Dim i As Long
    Dim sScript As String
    Dim bChecked As Boolean
    With lvScripts
        With .SelectedItem
            iTmp = .Index
            sScript = .SubItems(1)
            bChecked = .Checked
        End With
        With .ListItems
            .Remove iTmp  ' remove selected item
                ' now insert before previous item
            With .Add(iTmp - 1, , MakeLabel(iTmp - 1))
                .SubItems(1) = sScript
                .Checked = bChecked
            End With
        End With
        For i = iTmp To .ListItems.Count
            .ListItems(i).Text = MakeLabel(i)
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
' Date-Time  :       03/01/2004-21:45:43
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdUpdate_Click()
    Dim sScript As String
    sScript = Me.txtScript
    If Len(sScript) > 0 Then
        lvScripts.SelectedItem.SubItems(1) = sScript
'        PutMyString IniSectionText, IniTextScriptPrefix & lvScripts.SelectedItem.Text, sScript
        lvScripts.SelectedItem.SubItems(2) = txtScriptExts
        cmdOK.Enabled = True
        cmdWizard.Enabled = True
        bInsertInProgress = False
        CheckSelection
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        ShowHelpTopic HH_Text_Input_Auto_Recognition_Window
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
' Date-Time  :       03/01/2004-00:43:16
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()
'    magic for me!
    If InIDE(Me.hWnd) Then
        cmdAuto.Enabled = True
        cmdAuto.Visible = True
    End If
    
    LocaliseForm Me, 1300
    cmdOK.Enabled = False
    gsClose = LoadResStringL(1316)
    gsCancel = LoadResStringL(1301)
    
    LoadScriptList
    
' this must be last as the caption gets set to Cancel in StartEdit
    cmdCancel.Caption = gsClose
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       lvScripts_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       03/01/2004-22:08:01
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub lvScripts_Click()
    CheckSelection
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       lvScripts_ItemCheck
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       10/02/2005-11:45:12
'
' Parameters :       Item (MSComctlLib.ListItem)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub lvScripts_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    cmdOK.Enabled = True
    cmdCancel.Caption = gsCancel
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtScriptExts_Change
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       11/03/2005-11:42:37
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtScriptExts_Change()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtScriptExts_LostFocus
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       11/03/2005-12:01:03
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtScriptExts_LostFocus()
    Dim sTmp As String
    Dim vExts As Variant
    Dim sNew As String
    If Len(txtScriptExts) = 0 Then
        Exit Sub
    End If
    vExts = Split(txtScriptExts, ",")
    Dim i As Integer
    For i = LBound(vExts) To UBound(vExts)
        sTmp = UCase$(CStr(vExts(i)))
        If Left(sTmp, 1) = "." Then
            sTmp = Mid$(sTmp, 2)
        End If
        If Len(sNew) = 0 Then
            sNew = sTmp
        Else
            sNew = sNew & "," & sTmp
        End If
    Next
    If sNew <> txtScriptExts Then
        txtScriptExts = sNew
    End If
End Sub

Private Sub LoadScriptList()
    Dim iScript As Long
    Dim vArr As Variant
    Dim sTmp As String
    Dim i As Integer
    Dim vExtList As Variant
    Dim iExt As Integer
    Dim sScript As String
    Dim sExt As String
    Dim sExtList As String

    Me.lvScripts.ListItems.Clear
    vExtList = GetIniKeyList(IniSectionTextExtension)
    
    iScript = 1
nextscript:
    sTmp = GetMyString(IniSectionText, _
        IniTextScriptPrefix & CStr(iScript), "")
    If sTmp = "" Then GoTo donescripts
    With Me.lvScripts.ListItems.Add(iScript, , MakeLabel(iScript))
        vArr = Split(sTmp, ",")
        If UBound(vArr) > 0 Then
            If IsNumeric(vArr(0)) Then
                .Checked = (CLng(vArr(0)) <> 0)
            Else
                .Checked = False
            End If
            .SubItems(1) = vArr(1)
        Else
            .Checked = True
            .SubItems(1) = sTmp
        End If
        sExtList = ""
        For iExt = LBound(vExtList) To UBound(vExtList)
            sExt = CStr(vExtList(iExt))
            sScript = GetMyString(IniSectionTextExtension, sExt, "")
            If sScript = .SubItems(1) Then
                If Len(sExtList) = 0 Then
                    sExtList = sExt
                Else
                    sExtList = sExtList & "," & sExt
                End If
            End If
            .SubItems(2) = sExtList
        Next
    End With
gonext:
    iScript = iScript + 1
    GoTo nextscript
donescripts:
    iOriginalCount = iScript - 1
    CheckSelection
End Sub
