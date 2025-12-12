VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processing..."
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Left            =   4440
      Top             =   840
   End
   Begin MSComctlLib.ProgressBar pbProgress 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Now processing:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      WhatsThisHelpID =   1051
      Width           =   1695
   End
   Begin VB.Label lblPosition 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label lblPhase 
      Caption         =   "Show Phase Here..."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : frmProgress
'    Project    : MT2OFX
'
'    Description: Progress Bar dialog (modeless)
'
'    Modified   : $Author: Colin $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/frmProgress.frm 7     19/04/08 22:08 Colin $"
' $History: frmProgress.frm $
' 
' *****************  Version 7  *****************
' User: Colin        Date: 19/04/08   Time: 22:08
' Updated in $/MT2OFX
'
' *****************  Version 6  *****************
' User: Colin        Date: 7/12/06    Time: 14:59
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2

'</CSCC>
Public lMax As Long
Public lMin As Long
Public Factor As Long

Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Integer

Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Private bIsTopmost As Boolean

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Form_Load
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       31/05/2005-22:44:32
'
' Parameters :
' 20061015 CS: delay is now configurable as [General]ProgressBarDelay. Default is 1500. If the delay
' is set to 0, the timer will not be enabled and will therefore not fire; the progress dialog will remain
' invisible and non-topmost.
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()
    Dim iDelay As Long
    iDelay = Cfg.ProgressBarDelay
    If iDelay < 0 Or iDelay > 65535 Then
        iDelay = 0
    End If
    LocaliseForm Me, 1050
    lblPosition = ""
    If iDelay > 0 Then
        Timer1.Interval = iDelay
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
    End If
    ResetProgress
    bIsTopmost = False
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Timer1_Timer
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       31/05/2005-22:44:05
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Timer1_Timer()
    On Error Resume Next
    Me.Visible = True
    SetTopmost True
    Timer1.Enabled = False
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       SetProgress
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       31/05/2005-22:57:39
'
' Parameters :       lVal (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub SetProgress(lVal As Long)
    On Error Resume Next
'    Debug.Print "Min:" & lMin & " Max:" & lMax & " Val:" & lVal
    If lVal >= lMax Then
        pbProgress.Value = pbProgress.Max
    ElseIf lVal <= lMin Then
        pbProgress.Value = pbProgress.Min
    Else
        pbProgress.Value = (Factor * (lVal - lMin)) / (lMax - lMin)
    End If
    lblPosition.Caption = CStr(lVal) & "/" & CStr(lMax)
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ResetProgress
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       31/05/2005-22:59:40
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub ResetProgress()
    Factor = 1000
    pbProgress.Value = 0
    pbProgress.Min = 0
    pbProgress.Max = Factor
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       SetTopmost
' Description:       Manipulate "topmost" setting, returning old value
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       8/15/2006-16:40:22
'
' Parameters :       bTopmost (Boolean)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function SetTopmost(bTopmost As Boolean) As Boolean
    SetTopmost = bIsTopmost
    If bTopmost Then
        SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags
    Else
        SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags
    End If
    bIsTopmost = bTopmost
End Function

