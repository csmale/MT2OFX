VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MT2OFX Options"
   ClientHeight    =   8775
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   12105
   ControlBox      =   0   'False
   HelpContextID   =   1200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCustomOutput 
      Caption         =   "Custom Output Formats"
      Height          =   375
      Left            =   1200
      TabIndex        =   60
      Top             =   5760
      Width           =   3375
   End
   Begin VB.CommandButton cmdViewLog 
      Caption         =   "View Log File"
      Height          =   375
      Left            =   1200
      TabIndex        =   59
      Top             =   5280
      WhatsThisHelpID =   1263
      Width           =   3375
   End
   Begin VB.Frame FraProgramOptions 
      Caption         =   "Program Options"
      Height          =   3615
      Left            =   120
      TabIndex        =   44
      Top             =   120
      Width           =   5415
      Begin VB.CheckBox chkNoConfirmImport 
         Caption         =   "No confirmation before importing"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   3240
         WhatsThisHelpID =   1261
         Width           =   5175
      End
      Begin VB.TextBox txtScriptEdit 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CommandButton cmdBrowse 
         Height          =   375
         Left            =   4920
         Picture         =   "frmOptions.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtMT940Exts 
         Height          =   285
         Left            =   2640
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtTextExts 
         Height          =   285
         Left            =   2640
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox txtDefaultPayee 
         Height          =   285
         Left            =   2640
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2880
         Width           =   2655
      End
      Begin VB.CheckBox chkAutoBrowse 
         Caption         =   "Show Open Dialog on startup"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "If checked the output will be imported directly into Money when conversion is complete."
         Top             =   240
         WhatsThisHelpID =   1201
         Width           =   4095
      End
      Begin VB.ComboBox cbOutputTo 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   3855
      End
      Begin VB.CheckBox chkSaveClipOutput 
         Caption         =   "Save Output from Clipboard conversion"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         WhatsThisHelpID =   1229
         Width           =   5295
      End
      Begin VB.CommandButton cmdOutProgs 
         Caption         =   "Post-conversion progs"
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   1320
         WhatsThisHelpID =   1202
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Script Editor"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1800
         WhatsThisHelpID =   1218
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "MT940 File Extensions"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   2160
         WhatsThisHelpID =   1214
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Text File Extensions"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   2520
         WhatsThisHelpID =   1215
         Width           =   2535
      End
      Begin VB.Label Label10 
         Caption         =   "Default Payee Name"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   2880
         WhatsThisHelpID =   1221
         Width           =   2655
      End
      Begin VB.Label Label13 
         Caption         =   "Save output:"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   600
         WhatsThisHelpID =   1203
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Output Options"
      Height          =   3855
      Left            =   5640
      TabIndex        =   38
      Top             =   120
      WhatsThisHelpID =   1264
      Width           =   6375
      Begin VB.ComboBox cbLineEnding 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CheckBox chkIncludeBOM 
         Caption         =   "Include Byte Order Marks in Unicode"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3480
         WhatsThisHelpID =   1265
         Width           =   6015
      End
      Begin VB.ComboBox cbCodePage 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2760
         Width           =   4095
      End
      Begin VB.CheckBox chkCompressSpaces 
         Caption         =   "Compress multiple spaces in payee and memo"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "If checked the output will be imported directly into Money when conversion is complete."
         Top             =   240
         WhatsThisHelpID =   1204
         Width           =   5295
      End
      Begin VB.ComboBox cbBookDateSelector 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Money only uses the Book Date. Select here which date you would like to see in Money."
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox cbOutputFileType 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtMemoSeparator 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         MaxLength       =   1
         TabIndex        =   17
         Text            =   ";"
         Top             =   2040
         Width           =   375
      End
      Begin VB.ComboBox cbPayeeCase 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2400
         Width           =   2415
      End
      Begin VB.CheckBox chkGenCheckNum 
         Caption         =   "Always generate cheque number"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         WhatsThisHelpID =   1220
         Width           =   5415
      End
      Begin VB.CheckBox chkNoSortTxns 
         Caption         =   "Do not sort transactions into date order"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         WhatsThisHelpID =   1231
         Width           =   5415
      End
      Begin VB.CheckBox chkSuppressEmpty 
         Caption         =   "Suppress empty statements"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         WhatsThisHelpID =   1254
         Width           =   5415
      End
      Begin VB.Label lblLineEnding 
         Caption         =   "Line Ending"
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   3120
         WhatsThisHelpID =   1266
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "Output Code Page"
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   2760
         WhatsThisHelpID =   1268
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "will be used as book date"
         Height          =   255
         Left            =   2400
         TabIndex        =   42
         Top             =   1320
         WhatsThisHelpID =   1205
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Default output file type"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1680
         WhatsThisHelpID =   1206
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Memo separator character"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   2040
         WhatsThisHelpID =   1223
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Payee case"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2400
         WhatsThisHelpID =   1208
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "OFX/QFX Output"
      Height          =   2055
      Left            =   5640
      TabIndex        =   36
      Top             =   4080
      WhatsThisHelpID =   1259
      Width           =   6375
      Begin VB.CheckBox chkForceBalDate 
         Caption         =   "Use last txn for missing balance date"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1440
         WhatsThisHelpID =   1262
         Width           =   6135
      End
      Begin VB.CheckBox chkUseOldFITID 
         Caption         =   "Use pre-3.5 FITID generation"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         ToolTipText     =   "If checked the output will be imported directly into Money when conversion is complete."
         Top             =   1680
         WhatsThisHelpID =   1256
         Width           =   6135
      End
      Begin VB.ComboBox cbQfxForceCurrency 
         Height          =   315
         Left            =   3000
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   960
         Width           =   2415
      End
      Begin VB.ComboBox cbOfxDecimal 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   600
         Width           =   2415
      End
      Begin VB.ComboBox cbOFXVersion 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Force currency in QFX output to"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   1080
         WhatsThisHelpID =   1257
         Width           =   2775
      End
      Begin VB.Label Label16 
         Caption         =   "Decimal point"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   720
         WhatsThisHelpID =   1255
         Width           =   2775
      End
      Begin VB.Label Label11 
         Caption         =   "OFX Version"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   360
         WhatsThisHelpID =   1222
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "QIF Output"
      Height          =   2415
      Left            =   5640
      TabIndex        =   31
      Top             =   6240
      WhatsThisHelpID =   1267
      Width           =   6375
      Begin VB.CheckBox chkQIFOutputUAmount 
         Caption         =   "Output Amount as U-line"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   2040
         WhatsThisHelpID =   1260
         Width           =   5775
      End
      Begin VB.ComboBox cbQifDate 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox cbQifSeparator 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CheckBox chkQIFNoAcctHdr 
         Caption         =   "No !Account header in QIF"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1800
         WhatsThisHelpID =   1225
         Width           =   5895
      End
      Begin VB.TextBox txtMaxMemoLength 
         Height          =   285
         Left            =   4800
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtQIFCustomDateFormat 
         Height          =   285
         Left            =   3000
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "QIF Custom Date Format"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   720
         WhatsThisHelpID =   1232
         Width           =   2775
      End
      Begin VB.Label Label7 
         Caption         =   "QIF Decimal Separator"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         WhatsThisHelpID =   1224
         Width           =   2775
      End
      Begin VB.Label Label12 
         Caption         =   "QIF Max Memo Length"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1440
         WhatsThisHelpID =   1230
         Width           =   5055
      End
      Begin VB.Label Label14 
         Caption         =   "QIF Date Format"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   360
         WhatsThisHelpID =   1212
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdPayee 
      Caption         =   "Payee Replacement..."
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   3840
      WhatsThisHelpID =   1217
      Width           =   3375
   End
   Begin MSComDlg.CommonDialog cdScriptEditor 
      Left            =   4560
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdText 
      Caption         =   "Scripts..."
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   4320
      WhatsThisHelpID =   1213
      Width           =   3375
   End
   Begin VB.CommandButton cmdBanks 
      Caption         =   "MT940..."
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   4800
      WhatsThisHelpID =   1209
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   29
      Top             =   7680
      WhatsThisHelpID =   1210
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   30
      Top             =   7680
      WhatsThisHelpID =   1211
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : frmOptions
'    Project    : MT2OFX
'
'    Description:
'
'    Modified   : $Author: Colin $ $Date: 15/11/10 0:00 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/frmOptions.frm 32    15/11/10 0:00 Colin $"
' $History: frmOptions.frm $
' 
' *****************  Version 32  *****************
' User: Colin        Date: 15/11/10   Time: 0:00
' Updated in $/MT2OFX
'
' *****************  Version 31  *****************
' User: Colin        Date: 19/12/09   Time: 1:09
' Updated in $/MT2OFX
' view log file
'
' *****************  Version 30  *****************
' User: Colin        Date: 24/11/09   Time: 22:04
' Updated in $/MT2OFX
' for 3.6 beta
'
' *****************  Version 29  *****************
' User: Colin        Date: 15/06/09   Time: 19:24
' Updated in $/MT2OFX
' For transfer to new laptop
'
' *****************  Version 28  *****************
' User: Colin        Date: 25/11/08   Time: 22:21
' Updated in $/MT2OFX
' moving vss server!
'
' *****************  Version 26  *****************
' User: Colin        Date: 19/04/08   Time: 23:12
' Updated in $/MT2OFX
' new style
' added codepage stuff
'
' *****************  Version 25  *****************
' User: Colin        Date: 7/12/06    Time: 13:21
' Updated in $/MT2OFX
' Added SkipEmptyStatements and SaveClipboardOutput
'
' *****************  Version 22  *****************
' User: Colin        Date: 2/11/05    Time: 23:03
' Updated in $/MT2OFX
' V3.4 beta 1
'
' *****************  Version 21  *****************
' User: Colin        Date: 8/05/05    Time: 12:43
' Updated in $/MT2OFX
' V3.3.8
'
' *****************  Version 20  *****************
' User: Colin        Date: 18/03/05   Time: 21:57
' Updated in $/MT2OFX
'
' *****************  Version 19  *****************
' User: Colin        Date: 6/03/05    Time: 23:41
' Updated in $/MT2OFX
'</CSCC>

Public OptCfg As ProgramConfig

Private gsClose As String
Private gsCancel As String
Private gsQFXUseOrigCurrency As String
Private aCurrencyList() As String

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       StartEdit
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       12/12/2003-00:35:28
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub StartEdit()
    OptCfg.Changed = True
    Me.cmdCancel.Caption = gsCancel
    Me.cmdOK.Enabled = True
End Sub
Private Sub cbBookDateSelector_Click()
    StartEdit
End Sub

Private Sub cbCodePage_Click()
    StartEdit
End Sub

Private Sub cbLineEnding_Click()
    StartEdit
End Sub

Private Sub cbOfxDecimal_Click()
    StartEdit
End Sub

Private Sub cbOFXVersion_Click()
    StartEdit
End Sub

Private Sub cbOutputFileType_Click()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cbOutputTo_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       09/10/2004-21:44:20
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cbOutputTo_Click()
    StartEdit
'    If cbOutputTo.ListIndex = OUTFILE_TEMP Then
'        chkAutoStartImport.Value = vbChecked
'        chkAutoStartImport.Enabled = False
'    Else
'        chkAutoStartImport.Enabled = True
'    End If
End Sub

Private Sub cbPayeeCase_Click()
    StartEdit
End Sub

Private Sub cbQfxForceCurrency_Click()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cbQifDate_Click
' Description:       handle click on QifDate
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       12/12/2003-21:29:09
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cbQifDate_Click()
    StartEdit
    On Error Resume Next
    txtQIFCustomDateFormat.Enabled = (cbQifDate.ListIndex = DATEFMT_CUSTOM)
    If txtQIFCustomDateFormat.Enabled Then
        txtQIFCustomDateFormat.SetFocus
        txtQIFCustomDateFormat.SelStart = 0
        txtQIFCustomDateFormat.SelLength = 9999
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cbQifSeparator_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       19/01/2004-23:02:50
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cbQifSeparator_Click()
    StartEdit
End Sub

Private Sub chkAutoBrowse_Click()
    StartEdit
End Sub

Private Sub chkAutoStartImport_Click()
    StartEdit
End Sub

Private Sub chkCompressSpaces_Click()
    StartEdit
End Sub

Private Sub chkForceBalDate_Click()
    StartEdit
End Sub

Private Sub chkGenCheckNum_Click()
    StartEdit
End Sub

Private Sub chkIncludeBOM_Click()
    StartEdit
End Sub

Private Sub chkNoConfirmImport_Click()
    StartEdit
End Sub

Private Sub chkNoSortTxns_Click()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       chkQIFNoAcctHdr_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       07/10/2004-23:47:42
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub chkQIFNoAcctHdr_Click()
    StartEdit
End Sub

Private Sub chkQIFOutputUAmount_Click()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       chkSaveClipOutput_Click
' Description:
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       11/29/2006-21:51:04
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub chkSaveClipOutput_Click()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       chkSuppressEmpty_Click
' Description:
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       8/25/2006-14:21:56
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub chkSuppressEmpty_Click()
    StartEdit
End Sub

Private Sub chkUseOldFITID_Click()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdBanks_Click
' Description:       handle click on Banks button
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       11/12/2003-21:56:00
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdBanks_Click()
    frmBankOptions.Show vbModal, Me
    If Not Me.cmdOK.Enabled Then
        Me.cmdCancel.Caption = gsClose
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdBrowse_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-22:07:21
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdBrowse_Click()
    Dim iErr As Long
    Dim iLen As Long
    Dim sTmp As String
        With cdScriptEditor
        .DialogTitle = LoadResStringL(1218)
        sTmp = String(255, vbNullChar)
        iErr = SHGetFolderPath(0, CSIDL_PROGRAM_FILES, 0, SHGFP_TYPE_CURRENT, sTmp)
        If iErr = 0 Then
            ChDir Left$(sTmp, InStr(sTmp, vbNullChar))
            .InitDir = Left$(sTmp, InStr(sTmp, vbNullChar))
        End If
        .FileName = ""
        .Filter = LoadResStringL(1219)
        .FilterIndex = 1
        .Flags = cdlOFNFileMustExist + cdlOFNLongNames + cdlOFNExplorer _
            + cdlOFNHideReadOnly + cdlOFNPathMustExist
        .ShowOpen
        If .FileName <> "" Then
            Me.txtScriptEdit = .FileName
            StartEdit
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdCustomOutput_Click
' Description:       Handle click on Custom Output Formats
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       18/04/2010-22:38:40
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdCustomOutput_Click()
    Dim f As New frmCustomOutput
    With f
        .ScriptDir = OptCfg.CustomOutputPath
        .Show vbModal
    End With
End Sub

Private Sub cmdOK_Click()
    If OptCfg.Changed Then
        ReadControls
        If Not OptCfg.Save() Then
            MyMsgBox LoadResStringL(109)
        End If
    End If
    Unload Me
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdPayee_Click
' Description:       Handle click on "Payees..."
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       05/07/2004-20:26:43
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdPayee_Click()
    DebugLog "Starting Payee Replacement", logDEBUG
    frmPayee.Show vbModal, Me
    DebugLog "Finished Payee Replacement", logDEBUG
    If Not Me.cmdOK.Enabled Then
        Me.cmdCancel.Caption = gsClose
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdText_Click
' Description:       handle click on Text... button
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-17:43:10
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdText_Click()
    frmAutoInput.Show vbModal, Me
    If Not Me.cmdOK.Enabled Then
        Me.cmdCancel.Caption = gsClose
    End If
End Sub

Private Sub cmdOutProgs_Click()
    Set frmOutProg.OptCfg = OptCfg
    frmOutProg.Show vbModal, Me
    cmdOK.Enabled = OptCfg.Changed
    If Not cmdOK.Enabled Then
        Me.cmdCancel.Caption = gsClose
    End If
End Sub

Private Sub cmdViewLog_Click()
    Dim sTmp As String
    If PathIsRelative(Cfg.LogFile) Then
        sTmp = Cfg.AppDataPath & "\" & Cfg.LogFile
    Else
        sTmp = Cfg.LogFile
    End If
    DoDefaultFileAction sTmp
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        ShowHelpTopic HH_Options_Window
        KeyCode = 0
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim iTmp As Long
    Dim sTmp As String, sTmp2 As String
    Dim lCP() As Long
    Dim xChoice As Variant
    
    If OptCfg Is Nothing Then
        Debug.Assert False
        Exit Sub
    End If
    If Not InIDE(Me.hWnd) Then
        Me.cmdCustomOutput.TabStop = False
        Me.cmdCustomOutput.Visible = False
        Me.cmdCustomOutput.Enabled = False
    End If
    LocaliseForm Me
    gsClose = LoadResStringL(1216)
    gsCancel = LoadResStringL(1210)
    gsQFXUseOrigCurrency = LoadResStringL(1258)
    With Me.cbBookDateSelector
        .Clear
' bdm*Date must be in ascending order and start at 0 - see ProgramConfig.cls
        .AddItem LoadResStringL(105), bdmBookDate
        .AddItem LoadResStringL(106), bdmValueDate
        .AddItem LoadResStringL(107), bdmTransDate
    End With
    With Me.cbOutputFileType
        .Clear
        .AddItem "OFX"
        .AddItem "OFC"
        .AddItem "QIF"
        .AddItem "QFX"
    End With
    With Me.cbOutputTo
        .Clear
        .AddItem LoadResStringL(1226), OUTFILE_AUTO
        .AddItem LoadResStringL(1227), OUTFILE_ASK
        .AddItem LoadResStringL(1228), OUTFILE_TEMP
    End With
    With Me.cbOFXVersion
        .Clear
        .AddItem "102"
        .AddItem "151"
        .AddItem "160"
        .AddItem "200"
        .AddItem "201"
' 20041201 CS Added 2.0.2
        .AddItem "202"
' 20090417 CS Added 2.0.3
        .AddItem "203"
' 20080104 CS Added 2.1.0
        .AddItem "210"
' 20090417 CS Added 2.1.1
        .AddItem "211"
    End With
    ' 0=No Change (default), 1=Upper Case, 2=Lower Case, 3=Proper Case
    With Me.cbPayeeCase
        .Clear
        .AddItem LoadResStringL(120), 0
        .AddItem LoadResStringL(121), 1
        .AddItem LoadResStringL(122), 2
        .AddItem LoadResStringL(123), 3
    End With
    With Me.cbQifDate
        .Clear
        .AddItem LoadResStringL(125), DATEFMT_MDY
        .AddItem LoadResStringL(124), DATEFMT_DMY
        .AddItem LoadResStringL(133), DATEFMT_YMD
        .AddItem LoadResStringL(134), DATEFMT_SYSTEM
        .AddItem LoadResStringL(139), DATEFMT_CUSTOM
    End With
    Me.txtQIFCustomDateFormat = OptCfg.QifCustomDateFormat
    With Me.cbQifSeparator
        .Clear
        .AddItem LoadResStringL(137), 0
        .AddItem LoadResStringL(138), 1
        .AddItem LoadResStringL(134), 2
    End With
    Me.cbOutputTo.ListIndex = OptCfg.PromptForOutput
    Me.cbBookDateSelector.ListIndex = OptCfg.BookDateMode
'    Me.chkAutoStartImport.Value = IIf(OptCfg.AutoStartImport, vbChecked, vbUnchecked)
    Me.chkCompressSpaces.Value = IIf(OptCfg.CompressSpaces, vbChecked, vbUnchecked)
    Me.cbOutputFileType = OptCfg.OutputFileType
    Me.cbOFXVersion = CStr(OptCfg.OFXVersion)
    With Me.cbOfxDecimal
        .Clear
        .AddItem LoadResStringL(137), 0
        .AddItem LoadResStringL(138), 1
    End With
    cbOfxDecimal.ListIndex = FindInCombo(cbOfxDecimal, OptCfg.OFXDecimal)
    Me.chkForceBalDate.Value = IIf(OptCfg.OFXForceRealBalanceDate, vbChecked, vbUnchecked)
    Me.cbPayeeCase.ListIndex = OptCfg.PayeeCase
    Me.chkAutoBrowse.Value = IIf(OptCfg.AutoBrowse, vbChecked, vbUnchecked)
    Me.txtMemoSeparator = OptCfg.MemoDelimiter
    Me.cbQifDate.ListIndex = OptCfg.QifDateFormat
    If OptCfg.QifDecimalSeparator = "" Then
        Me.cbQifSeparator.ListIndex = 2
    Else
        Me.cbQifSeparator.ListIndex = FindInCombo(Me.cbQifSeparator, OptCfg.QifDecimalSeparator)
    End If
    Me.chkQIFNoAcctHdr.Value = IIf(OptCfg.QifNoAcctHeader, vbChecked, vbUnchecked)
    Me.chkQIFOutputUAmount.Value = IIf(OptCfg.QifOutputUAmount, vbChecked, vbUnchecked)
    Me.txtMaxMemoLength = CStr(OptCfg.QifMaxMemoLength)
    Me.txtScriptEdit = OptCfg.ScriptEditor
    Me.txtMT940Exts = OptCfg.AutoMT940
    Me.txtTextExts = OptCfg.AutoText
    Me.txtDefaultPayee = OptCfg.DefaultPayee
    Me.chkGenCheckNum.Value = IIf(OptCfg.GenerateCheckNum, vbChecked, vbUnchecked)
    Me.chkNoSortTxns.Value = IIf(OptCfg.NoSortTxns, vbChecked, vbUnchecked)
    Me.chkSuppressEmpty = IIf(OptCfg.SuppressEmptyStatements, vbChecked, vbUnchecked)
    Me.chkSaveClipOutput = IIf(OptCfg.SaveClipboardOutput, vbChecked, vbUnchecked)
    
    lCP = GetCodePageList()
    cbCodePage.Clear
    cbCodePage.AddItem "      " & LoadResStringL(134)
    For i = 0 To UBound(lCP)
        iTmp = lCP(i)
        If IsValidCodePage(iTmp) And iTmp <> CP_USER Then
            sTmp = GetCodePageName(iTmp)
            If Left$(sTmp, 1) <> "_" Then
                sTmp2 = Right$("     " & CStr(iTmp), 5)
                cbCodePage.AddItem sTmp2 & " : " & sTmp
            End If
        End If
    Next
    If OptCfg.OutputCodePage = CP_ACP Then
        sTmp = "     "
    Else
        sTmp = Right$("     " & CStr(OptCfg.OutputCodePage), 5)
    End If
    cbCodePage.ListIndex = FindInCombo(cbCodePage, sTmp)
    Me.chkIncludeBOM = IIf(OptCfg.OutputBOM, vbChecked, vbUnchecked)

    iTmp = -1
    With Me.cbLineEnding
        .AddItem "LF":      .ItemData(.newIndex) = leLF:   If OptCfg.LineEnding = leLF Then iTmp = .newIndex
        .AddItem "CR":      .ItemData(.newIndex) = leCR:   If OptCfg.LineEnding = leCR Then iTmp = .newIndex
        .AddItem "CR+LF":   .ItemData(.newIndex) = leCRLF: If OptCfg.LineEnding = leCRLF Then iTmp = .newIndex
        .AddItem "LF+CR":   .ItemData(.newIndex) = leLFCR: If OptCfg.LineEnding = leLFCR Then iTmp = .newIndex
        .ListIndex = iTmp
    End With
    
    If IsXpOrLater() Then
        Me.chkUseOldFITID.Enabled = True
        Me.chkUseOldFITID = IIf(OptCfg.UseOldFITID, vbChecked, vbUnchecked)
    Else
        Me.chkUseOldFITID = vbChecked
        Me.chkUseOldFITID.Enabled = False
    End If

    aCurrencyList = GetCurrencyList
    cbQfxForceCurrency.Clear
    cbQfxForceCurrency.AddItem gsQFXUseOrigCurrency
    For Each xChoice In aCurrencyList
        If FindInCombo(cbQfxForceCurrency, CStr(xChoice)) = -1 Then
            cbQfxForceCurrency.AddItem CStr(xChoice)
        End If
    Next
    If Len(OptCfg.QFXForceCurrency) > 0 Then
        cbQfxForceCurrency.ListIndex = FindInCombo(cbQfxForceCurrency, OptCfg.QFXForceCurrency)
    Else
        cbQfxForceCurrency.ListIndex = FindInCombo(cbQfxForceCurrency, gsQFXUseOrigCurrency)
    End If
    chkNoConfirmImport = IIf(OptCfg.NoConfirmImport, vbChecked, vbUnchecked)
    
    Me.cmdCancel.Caption = gsClose
    Me.cmdOK.Enabled = False
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ReadControls
' Description:       Read the form controls into the config object
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       12/12/2003-00:18:46
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub ReadControls()
    Dim sTmp As String
    Dim iTmp As Long
    OptCfg.PromptForOutput = Me.cbOutputTo.ListIndex
    OptCfg.BookDateMode = Me.cbBookDateSelector.ListIndex
    OptCfg.OutputFileType = Me.cbOutputFileType
    OptCfg.OFXVersion = CInt(Me.cbOFXVersion)
    OptCfg.PayeeCase = Me.cbPayeeCase.ListIndex
    OptCfg.AutoBrowse = (Me.chkAutoBrowse.Value = vbChecked)
'    OptCfg.AutoStartImport = (Me.chkAutoStartImport.Value = vbChecked)
    OptCfg.MemoDelimiter = Me.txtMemoSeparator
    OptCfg.QifDateFormat = Me.cbQifDate.ListIndex
    OptCfg.QifCustomDateFormat = Me.txtQIFCustomDateFormat
    If Me.cbQifSeparator.ListIndex = 2 Then
        OptCfg.QifDecimalSeparator = ""
    Else
        OptCfg.QifDecimalSeparator = Left$(Me.cbQifSeparator, 1)
    End If
    OptCfg.QifNoAcctHeader = (Me.chkQIFNoAcctHdr.Value = vbChecked)
    OptCfg.QifMaxMemoLength = CLng(Me.txtMaxMemoLength)
    OptCfg.QifOutputUAmount = (Me.chkQIFOutputUAmount.Value = vbChecked)
    OptCfg.ScriptEditor = Me.txtScriptEdit
    OptCfg.AutoMT940 = Me.txtMT940Exts
    OptCfg.AutoText = Me.txtTextExts
    OptCfg.DefaultPayee = Me.txtDefaultPayee
    OptCfg.GenerateCheckNum = (Me.chkGenCheckNum.Value = vbChecked)
    OptCfg.NoSortTxns = (Me.chkNoSortTxns.Value = vbChecked)
    OptCfg.SuppressEmptyStatements = (Me.chkSuppressEmpty.Value = vbChecked)
    OptCfg.SaveClipboardOutput = (Me.chkSaveClipOutput.Value = vbChecked)
    OptCfg.OutputBOM = (Me.chkIncludeBOM.Value = vbChecked)
    sTmp = Trim(Left$(cbCodePage, 5))
    If Len(sTmp) = 0 Then
        OptCfg.OutputCodePage = CP_ACP
    Else
        OptCfg.OutputCodePage = CLng(sTmp)
    End If
    OptCfg.LineEnding = cbLineEnding.ItemData(cbLineEnding.ListIndex)
    OptCfg.OFXDecimal = Left$(cbOfxDecimal, 1)
    OptCfg.OFXForceRealBalanceDate = (Me.chkForceBalDate.Value = vbChecked)
    OptCfg.UseOldFITID = (Me.chkUseOldFITID.Value = vbChecked)
    If Me.cbQfxForceCurrency = gsQFXUseOrigCurrency Then
        OptCfg.QFXForceCurrency = ""
    Else
        OptCfg.QFXForceCurrency = Left$(Me.cbQfxForceCurrency, 3)
    End If
    OptCfg.NoConfirmImport = (chkNoConfirmImport.Value = vbChecked)
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtDefaultPayee_Change
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       01/09/2004-22:01:32
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtDefaultPayee_Change()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtMaxMemoLength_Change
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       26-Jan-2005-22:07:32
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtMaxMemoLength_Change()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtMaxMemoLength_KeyPress
' Description:       make txtMaxMemoLength numbers only
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       26-Jan-2005-22:06:19
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtMaxMemoLength_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    'Allows only numeric keys and (backspace) and (-) keys
    'and removes any others by setting it to null = chr(0)
    Case 8, 48 To 57, 127
    Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtMemoSeparator_Change()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtMT940Exts_Change
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       01/02/2004-22:29:39
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtMT940Exts_Change()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtMT940Exts_LostFocus
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       01/02/2004-22:33:38
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtMT940Exts_LostFocus()
    Dim sTmp As String
    sTmp = UCase$(txtMT940Exts)
    If sTmp <> txtMT940Exts Then txtMT940Exts = sTmp
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtQIFCurrency_Change
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       07/10/2004-23:38:08
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtQIFCurrency_Change()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtQIFCustomDateFormat_Change
' Description:
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       8/31/2005-23:16:11
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtQIFCustomDateFormat_Change()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtTextExts_Change
' Description:       [type_description_here]
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       15/02/2004-23:53:01
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtTextExts_Change()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtTextExts_LostFocus
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       15/02/2004-23:53:26
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtTextExts_LostFocus()
    Dim sTmp As String
    sTmp = UCase$(txtTextExts)
    If sTmp <> txtTextExts Then txtTextExts = sTmp
End Sub
