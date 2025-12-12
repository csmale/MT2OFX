VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomOutput 
   Caption         =   "Custom Output Scripts"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   6480
      TabIndex        =   2
      Top             =   3000
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   7560
      TabIndex        =   1
      Top             =   3000
      Width           =   990
   End
   Begin MSComctlLib.ListView lvScripts 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Filter"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Default"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Script"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmCustomOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xScripts As New CustomOutputList
Public ScriptDir As String

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdOK_Click
' Description:       Click on OK button
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       18/04/2010-22:42:53
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdOK_Click()
    Me.Hide
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Form_Load
' Description:       Load event
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       18/04/2010-22:44:10
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()
    Me.lvScripts.ListItems.Clear
    If Len(ScriptDir) = 0 Then
        Debug.Assert Len(ScriptDir) <> 0
        Exit Sub
    End If
    Set xScripts = New CustomOutputList
    xScripts.Load ScriptDir
    Dim xFmt As CustomOutputFormat
    Dim xItem As ListItem
    For Each xFmt In xScripts
        Set xItem = lvScripts.ListItems.Add(, , xFmt.FormatName)
        With xItem
            .SubItems(1) = xFmt.Filters
            .SubItems(2) = xFmt.DefaultExtension
            .SubItems(3) = xFmt.ScriptName
        End With
    Next
End Sub
