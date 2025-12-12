VERSION 5.00
Begin VB.Form frmViewer 
   Caption         =   "Statement Data Viewer"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmStatement 
      Caption         =   "Statement Detail"
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   7335
      Begin VB.CommandButton cmdViewTxns 
         Caption         =   "View Transactions"
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label10 
         Caption         =   "Account:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Bank/Branch:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Available Balance:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Closing Balance:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Opening Balance:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Currency:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.ComboBox cbStatement 
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Frame frmSession 
      Caption         =   "Session"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7335
      Begin VB.TextBox txtServerTime 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   720
         Width           =   5895
      End
      Begin VB.TextBox txtOutputFile 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   480
         Width           =   5895
      End
      Begin VB.TextBox txtInputFile 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label5 
         Caption         =   "Server time:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Output file:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Input file:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Select statement:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : frmViewer
'    Project    : MT2OFX
'
'    Description: Viewer/editor for session/statement/transactions
'
'    Modified   : $Author: Colin $ $Date: 8/05/05 12:43 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/frmViewer.frm 2     8/05/05 12:43 Colin $"
' $History: frmViewer.frm $
' 
' *****************  Version 2  *****************
' User: Colin        Date: 8/05/05    Time: 12:43
' Updated in $/MT2OFX
' V3.3.8
'</CSCC>
