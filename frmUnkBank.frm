VERSION 5.00
Begin VB.Form frmUnkBank 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Unknown Bank"
   ClientHeight    =   2940
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   ControlBox      =   0   'False
   HelpContextID   =   1100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdMoreInfo 
      Caption         =   "More Info..."
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   2040
      WhatsThisHelpID =   1103
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      WhatsThisHelpID =   1104
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   $"frmUnkBank.frx":0000
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      WhatsThisHelpID =   1102
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "MT2OFX has been unable to identify the bank which produced the file ""%1""."
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      WhatsThisHelpID =   1101
      Width           =   4335
   End
End
Attribute VB_Name = "frmUnkBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : frmUnkBank
'    Project    : MT2OFX
'
'    Description:
'
'    Modified   : $Author: Colin $ $Date: 28/10/09 22:49 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/frmUnkBank.frm 9     28/10/09 22:49 Colin $"
' $History: frmUnkBank.frm $
' 
' *****************  Version 9  *****************
' User: Colin        Date: 28/10/09   Time: 22:49
' Updated in $/MT2OFX
' bit more room for the message
'
' *****************  Version 5  *****************
' User: Colin        Date: 6/03/05    Time: 23:41
' Updated in $/MT2OFX
'</CSCC>

Public FileName As String

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdMoreInfo_Click
' Description:       Show More Info about new banks
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       28/11/2003-22:55:09
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdMoreInfo_Click()
    ShowHelpTopic HH_Supporting_a_New_Bank
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdOK_Click
' Description:       OK button
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       28/11/2003-22:54:44
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdOK_Click()
    Unload Me
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Form_KeyDown
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/08/2004-22:15:59
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        ShowHelpTopic HH_Supporting_a_New_Bank
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
' Date-Time  :       28/11/2003-22:50:29
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()
    LocaliseForm Me
    Me.Label1.Caption = Replace(Me.Label1.Caption, "%1", FileName)
End Sub
