VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmMain 
   Caption         =   "MT940 to OFX"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9240
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInFile 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Name of input file in MT940 format"
      Top             =   120
      Width           =   7575
   End
   Begin VB.CommandButton cmdLanguage 
      Caption         =   "Language..."
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.Timer tmClipMon 
      Left            =   3240
      Top             =   840
   End
   Begin VB.CommandButton cmdCvtClip 
      Caption         =   "From Clipboard"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7800
      TabIndex        =   7
      ToolTipText     =   "Click here to convert the contents of the Windows Clipboard"
      Top             =   1080
      WhatsThisHelpID =   1012
      Width           =   1335
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Options..."
      Height          =   375
      HelpContextID   =   1
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Configure program operation"
      Top             =   600
      WhatsThisHelpID =   1011
      Width           =   1215
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   1800
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help..."
      Height          =   375
      HelpContextID   =   1
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Show program usage information"
      Top             =   1800
      WhatsThisHelpID =   1001
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About..."
      Height          =   375
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Show program version information"
      Top             =   1440
      WhatsThisHelpID =   1008
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      HelpContextID   =   3
      Left            =   7800
      TabIndex        =   3
      ToolTipText     =   "Click here to close the program"
      Top             =   1800
      WhatsThisHelpID =   1004
      Width           =   1335
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert!"
      Height          =   375
      Left            =   7800
      TabIndex        =   2
      ToolTipText     =   "Click here to carry out the conversion"
      Top             =   600
      WhatsThisHelpID =   1003
      Width           =   1335
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      ToolTipText     =   "Click here to select an input file"
      Top             =   120
      WhatsThisHelpID =   1002
      Width           =   1335
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl2 
      Left            =   1800
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : frmMain
'    Project    : MT2OFX
'
'    Description: Main Form
'
'    Modified   : $Author: Colin $ $Date: 15/11/10 0:08 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/frmMain.frm 32    15/11/10 0:08 Colin $"
' $History: frmMain.frm $
' 
' *****************  Version 32  *****************
' User: Colin        Date: 15/11/10   Time: 0:08
' Updated in $/MT2OFX
'
' *****************  Version 31  *****************
' User: Colin        Date: 5/01/10    Time: 17:56
' Updated in $/MT2OFX
' Fixed problem causing main form to restart when closing prog
'
' *****************  Version 30  *****************
' User: Colin        Date: 6/10/09    Time: 0:35
' Updated in $/MT2OFX
' added watcher support for explicit output type
'
' *****************  Version 29  *****************
' User: Colin        Date: 15/06/09   Time: 19:24
' Updated in $/MT2OFX
' For transfer to new laptop
'
' *****************  Version 29  *****************
' User: Colin        Date: 17/01/09   Time: 23:30
' Updated in $/MT2OFX
'
' *****************  Version 28  *****************
' User: Colin        Date: 25/11/08   Time: 22:21
' Updated in $/MT2OFX
' moving vss server!
'
' *****************  Version 26  *****************
' User: Colin        Date: 20/04/08   Time: 10:04
' Updated in $/MT2OFX
' For 3.5 beta 1
'
' *****************  Version 25  *****************
' User: Colin        Date: 7/12/06    Time: 14:58
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 22  *****************
' User: Colin        Date: 2/11/05    Time: 23:03
' Updated in $/MT2OFX
' V3.4 beta 1
'
' *****************  Version 21  *****************
' User: Colin        Date: 6/05/05    Time: 23:12
' Updated in $/MT2OFX
' Added Error and Timeout event handlers to ScriptControl
'
' *****************  Version 20  *****************
' User: Colin        Date: 6/03/05    Time: 0:35
' Updated in $/MT2OFX
'
' *****************  Version 19  *****************
' User: Colin        Date: 6/03/05    Time: 0:24
' Updated in $/MT2OFX
'</CSCC>

' shadow copy of selected file name - does not go through text box so Unicode-pure
Private sFileNameW As String

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdAbout_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       10/02/2005-22:09:27
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdAbout_Click()
    frmAbout.Show vbModal
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdBrowse_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       10/02/2005-22:09:32
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdBrowse_Click()
    Dim iSlash As Integer
    Dim sTmp As String
    Dim i As Integer
    Dim vTmp As Variant
    Dim sFilter As String
    Dim iFilterIndex As Integer
    Dim sInitDir As String
    Dim sTitle As String
    
    sTmp = LoadResStringL(101) & "|" ' MT940 Statement Files
    vTmp = Split(Cfg.AutoMT940, ",")
    If TypeName(vTmp) = "String()" Then
        For i = 0 To UBound(vTmp)
            If i > 0 Then
                sTmp = sTmp & ";"
            End If
            sTmp = sTmp & "*." & vTmp(i)
        Next
    Else
        sTmp = sTmp & "*.*"
    End If
    sTmp = sTmp & "|" & LoadResStringL(136) & "|"   ' text files
    vTmp = Split(Cfg.AutoText, ",")
    If TypeName(vTmp) = "String()" Then
        For i = 0 To UBound(vTmp)
            If i > 0 Then
                sTmp = sTmp & ";"
            End If
            sTmp = sTmp & "*." & vTmp(i)
        Next
    Else
        sTmp = sTmp & "*.*"
    End If
    sFilter = sTmp & "|" & LoadResStringL(135) & "|*.*"    ' add All Files
    If InStr(Cfg.AutoText, Cfg.LastInputExtension) > 0 Then
        iFilterIndex = 2    ' assume text file
    Else
        iFilterIndex = 1    ' assume mt940 file
    End If
    sInitDir = Cfg.LastInputDirectory
    sTitle = LoadResStringL(103)
    sTmp = GetInputFileName("", sInitDir, sFilter, iFilterIndex, sTitle, Me.hWnd)
    If Len(sTmp) > 0 Then
        DBCSLog sTmp, "Returned by GetInputFileName"
' update unicode-safe shadow AFTER text box due to change event
        Me.txtInFile = sTmp
        sFileNameW = sTmp
        DBCSLog Me.txtInFile, "Re-read from text box"
        iSlash = InStrRev(sTmp, "\")
        Cfg.SetLastInputDirectory Left$(sTmp, iSlash - 1)
        Cfg.SetLastInputExtension GetExtension(sTmp)
    End If
    
#If False Then
    With cdInput
        .Filter = sFilter
        .FilterIndex = iFilterIndex
        .Flags = cdlOFNHideReadOnly + cdlOFNExplorer + cdlOFNFileMustExist _
            + cdlOFNLongNames
        .DialogTitle = LoadResStringL(103)
        .FileName = ""
        .InitDir = sInitDir
        .ShowOpen
        If .FileName <> "" Then
            DBCSLog .FileName, "Returned by ShowOpen"
            Me.txtInFile = .FileName
            DBCSLog Me.txtInFile, "Re-read from text box"
            iSlash = InStrRev(.FileName, "\")
            Cfg.SetLastInputDirectory Left$(.FileName, iSlash - 1)
            Cfg.SetLastInputExtension GetExtension(.FileName)
        End If
    End With
#End If
    CloseDBCSLog
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdClose_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       10/02/2005-22:09:43
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdClose_Click()
' allows event loop to check .Visible without creating a new form instance
    Me.Hide
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdConvert_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       10/02/2005-22:09:51
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdConvert_Click()
    Dim sIn As String
    Dim sOut As String
    Dim sType As String
    Dim sTmp As String
    Dim iType As Long

' use unicode-safe shadow name unless it has been cleared by manual editing of the text box
    sIn = sFileNameW
    If Len(sIn) = 0 Then
        sIn = Me.txtInFile
    End If
    DBCSLog sIn, "Input file"
' the output file is named the same as the input file, with the extension
' changed to "OFX" or whatever
    sOut = GetOutputFile(sIn, Me.hWnd, sType, iType)
    DBCSLog sOut, "Output file"
    If sOut = "" Then
        Exit Sub
    End If
    If Process(sIn, sOut, iType, sType) Then
        DoImport sOut, Cfg.NoConfirmImport
' make sure we give the import prog enough time to get started
        If Cfg.PromptForOutput = OUTFILE_TEMP Then
            DoEvents
            Sleep Cfg.TempFileDelay
        End If
    End If
    If Cfg.PromptForOutput = OUTFILE_TEMP Then
        RemoveTempFile sOut
    End If
    CloseLogFile
    CloseDBCSLog
    Exit Sub
baleout:
    If Err <> cdlCancel Then
        MyMsgBox Err.Description
    End If
    If Cfg.PromptForOutput = OUTFILE_TEMP Then
        RemoveTempFile sOut
    End If
    CloseLogFile
    CloseDBCSLog
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdCvtClip_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       10/02/2005-22:10:00
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdCvtClip_Click()
    Dim sIn As String
    Dim sOut As String
    Dim sTmp As String
    Dim sType As String
    Dim iType As Long
    
    HTMLClipboardData = GetHTMLClipboard(True)
    
    sIn = CaptureTextClipboard()
    DBCSLog sIn, "Clipboard temp file"
    If sIn = "" Then
        Exit Sub
    End If
    
' the dodgy bit is that we don't know if it's a MT940 file or some other
' format!! What are we going to do about this??? For the moment, it's a text file
' only (no proffering around the MT940 providers).

' the output file is named the same as the input file, with the extension
' changed to "OFX" or whatever
    sOut = GetOutputFile("", Me.hWnd, sType, iType)
    DBCSLog sOut, "Output file"
    If sOut = "" Then
        Exit Sub
    End If
    If Process(sIn, sOut, iType, sType) Then
        DoImport sOut, Cfg.NoConfirmImport
' make sure we give the import prog enough time to get started
        If Cfg.PromptForOutput = OUTFILE_TEMP Then
            DoEvents
            Sleep Cfg.TempFileDelay
        End If
    End If
    RemoveTempFile sIn
    If Cfg.PromptForOutput = OUTFILE_TEMP Then
        RemoveTempFile sOut
    End If
    CloseDBCSLog
    Exit Sub
baleout:
    If Err <> cdlCancel Then
        MyMsgBox Err.Description
    End If
    RemoveTempFile sIn
    If Cfg.PromptForOutput = OUTFILE_TEMP Then
        RemoveTempFile sOut
    End If
    CloseDBCSLog
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdHelp_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       10/02/2005-22:10:11
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdHelp_Click()
' open HTML help to table of contents at the default topic
    ShowHelpContents
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdLanguage_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/02/2005-22:24:59
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdLanguage_Click()
    frmLanguage.Show vbModal
    LocaliseForm Me
    Me.Caption = LoadResStringLEx(1000, CStr(App.Major) _
        & "." & CStr(App.Minor) & "." & CStr(App.Revision))
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdOptions_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       10/02/2005-22:09:14
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdOptions_Click()
    Set frmOptions.OptCfg = Cfg
    frmOptions.Show vbModal
End Sub


'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Form_KeyDown
' Description:       handle key down - in particular F1 for help
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       04/08/2004-09:36:39
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        ShowHelpTopic HH_MT2OFX_Main_Window
        KeyCode = 0
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Form_Load
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       10/02/2005-22:09:04
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()
    App.HelpFile = Cfg.HelpPath & "\mt2ofx.chm"
    LocaliseForm Me
    Me.Caption = LoadResStringLEx(1000, CStr(App.Major) _
        & "." & CStr(App.Minor) & "." & CStr(App.Revision))
    cmdConvert.Enabled = False
    tmClipMon.Interval = 1000
    tmClipMon.Enabled = True
    If Cfg.AutoBrowse Then
        Me.Visible = True
        cmdBrowse_Click
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Form_Terminate
' Description:       Wrap up when main form terminates
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       11/2/2005-20:00:54
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Terminate()
    CleanUpMutex
End Sub

Private Sub ScriptControl1_Error()
    Dim sx As ScriptEnv
    Set sx = GetScriptEnv()
    If Not sx Is Nothing Then
        sx.Abort
    End If
End Sub

Private Sub ScriptControl1_Timeout()
    Dim sx As ScriptEnv
    Set sx = GetScriptEnv()
    If Not sx Is Nothing Then
        sx.Abort
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       tmClipMon_Timer
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       26 Nov 2004-22:43:58
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub tmClipMon_Timer()
    On Error Resume Next
    cmdCvtClip.Enabled = (Len(Clipboard.GetText) > 0)
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtInFile_Change
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       10/02/2005-22:08:54
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtInFile_Change()
    cmdConvert.Enabled = (Len(txtInFile) > 0)
    sFileNameW = ""
End Sub
