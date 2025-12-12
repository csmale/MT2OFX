VERSION 5.00
Begin VB.Form frmNewExt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New File Extension Map"
   ClientHeight    =   1065
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3240
   ControlBox      =   0   'False
   HelpContextID   =   1250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Left            =   600
      Top             =   600
   End
   Begin VB.TextBox txtExt 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      WhatsThisHelpID =   1252
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      WhatsThisHelpID =   1253
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Extension:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      WhatsThisHelpID =   1251
      Width           =   975
   End
End
Attribute VB_Name = "frmNewExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : frmNewExt
'    Project    : MT2OFX
'
'    Description: Add New Input Extension
'
'    Modified   : $Author: Colin $ $Date: 6/03/05 0:35 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/frmNewExt.frm 6     6/03/05 0:35 Colin $"
' $History: frmNewExt.frm $
' 
' *****************  Version 6  *****************
' User: Colin        Date: 6/03/05    Time: 0:35
' Updated in $/MT2OFX
'</CSCC>

Public sExt As String

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CheckSelection
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-23:00:54
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub CheckSelection()
    If Len(Me.txtExt) > 0 Then
        Me.cmdOK.Enabled = True
    Else
        Me.cmdOK.Enabled = False
    End If
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdCancel_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-23:00:31
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
' Machine    :       55662M8
' Date-Time  :       02/01/2004-22:59:48
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdOK_Click()
    Dim sTmp As String
    sTmp = GetMyString(IniSectionTextExtension, txtExt, "")
    If sTmp <> "" Then
        MyMsgBox "Extension """ & txtExt & """ is already mapped.", vbCritical + vbOKOnly
        Exit Sub
    End If
    sExt = txtExt
    Unload Me
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Form_KeyDown
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/08/2004-21:29:08
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        ShowHelpTopic HH_New_Text_Extension_Window
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
' Date-Time  :       02/01/2004-23:02:04
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()
    LocaliseForm Me
    txtExt = UCase$(sExt)
    Timer1.Interval = 1
    Timer1.Enabled = True
    sExt = ""
    CheckSelection
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Timer1_Timer
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-23:02:54
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Timer1_Timer()
    txtExt.SetFocus
    Timer1.Enabled = False
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtExt_Change
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-23:10:18
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtExt_Change()
    CheckSelection
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtExt_LostFocus
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       17/02/2004-00:18:00
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtExt_LostFocus()
    If Left$(txtExt, 1) = "." Then
        txtExt = Mid$(txtExt, 2)
    End If
End Sub
