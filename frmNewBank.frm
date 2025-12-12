VERSION 5.00
Begin VB.Form frmNewBank 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Bank"
   ClientHeight    =   1635
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   ControlBox      =   0   'False
   HelpContextID   =   1150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MT2OFX.Hyperlink Hyperlink1 
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   1200
      Width           =   2535
      _extentx        =   4471
      _extenty        =   450
      font            =   "frmNewBank.frx":0000
      hyperlink       =   "http://www.swift.com/biconline/"
   End
   Begin VB.Timer Timer1 
      Left            =   4200
      Top             =   120
   End
   Begin VB.TextBox txtBankName 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox txtBankCode 
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   600
      WhatsThisHelpID =   1155
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      WhatsThisHelpID =   1154
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "To search for your bank's BIC, you can use this website:"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      WhatsThisHelpID =   1153
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Bank Name"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   600
      WhatsThisHelpID =   1152
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Bank Code (BIC)"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      WhatsThisHelpID =   1151
      Width           =   1815
   End
End
Attribute VB_Name = "frmNewBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : frmNewBank
'    Project    : MT2OFX
'
'    Description: Add New MT940 Bank
'
'    Modified   : $Author: Colin $ $Date: 6/03/05 0:35 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/frmNewBank.frm 6     6/03/05 0:35 Colin $"
' $History: frmNewBank.frm $
' 
' *****************  Version 6  *****************
' User: Colin        Date: 6/03/05    Time: 0:35
' Updated in $/MT2OFX
'</CSCC>

Public sBankCode As String
Public sBankName As String

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CheckValues
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-00:21:43
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub CheckValues()
    If Len(txtBankCode) > 0 And Len(txtBankName) > 0 Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdCancel_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-00:19:10
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdCancel_Click()
    sBankCode = ""
    sBankName = ""
    Unload Me
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdOK_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-00:17:11
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdOK_Click()
    Dim bc As New BankConfig
    If bc.Load(txtBankCode) Then
        MyMsgBox LoadResStringL(1156), vbOKOnly + vbCritical, LoadResStringL(1157)
        Exit Sub
    End If
    sBankCode = txtBankCode
    sBankName = txtBankName
    Unload Me
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Form_KeyDown
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/08/2004-21:02:25
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        ShowHelpTopic HH_New_Bank_Window
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
' Date-Time  :       02/01/2004-00:18:32
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()
    LocaliseForm Me
    Me.txtBankCode = sBankCode
    Me.txtBankName = sBankName
    Timer1.Interval = 1
    Timer1.Enabled = True
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Timer1_Timer
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-00:43:34
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Me.txtBankCode.SetFocus
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtBankCode_Change
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-00:21:01
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtBankCode_Change()
    CheckValues
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtBankName_Change
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-00:21:19
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtBankName_Change()
    CheckValues
End Sub
