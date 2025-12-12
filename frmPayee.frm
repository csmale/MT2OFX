VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPayee 
   Caption         =   "Payee Mapping"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13785
   ControlBox      =   0   'False
   HelpContextID   =   1700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   13785
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "Pattern Helper"
      Height          =   5895
      Left            =   11760
      TabIndex        =   48
      Top             =   0
      WhatsThisHelpID =   1724
      Width           =   1935
      Begin VB.CheckBox chkCapture 
         Caption         =   "Capture"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   4560
         WhatsThisHelpID =   1741
         Width           =   1575
      End
      Begin VB.CommandButton cmdInsertChars 
         Caption         =   "Insert Text"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   5400
         WhatsThisHelpID =   1740
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Caption         =   "How Many?"
         Height          =   1935
         Left            =   0
         TabIndex        =   50
         Top             =   2640
         WhatsThisHelpID =   1733
         Width           =   1935
         Begin VB.OptionButton optOne 
            Caption         =   "One"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            WhatsThisHelpID =   1734
            Width           =   1575
         End
         Begin VB.OptionButton optZeroOrOne 
            Caption         =   "Zero or one"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   480
            WhatsThisHelpID =   1735
            Width           =   1695
         End
         Begin VB.OptionButton optZeroOrMore 
            Caption         =   "Zero or more"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   720
            WhatsThisHelpID =   1736
            Width           =   1695
         End
         Begin VB.OptionButton optOneOrMore 
            Caption         =   "One or more"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   960
            WhatsThisHelpID =   1737
            Width           =   1695
         End
         Begin VB.TextBox txtNumberFrom 
            Height          =   375
            Left            =   120
            TabIndex        =   34
            Text            =   "Text2"
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox txtNumberTo 
            Height          =   375
            Left            =   1080
            TabIndex        =   35
            Text            =   "Text2"
            Top             =   1440
            Width           =   375
         End
         Begin VB.OptionButton optBetween 
            Caption         =   "Between:"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1200
            WhatsThisHelpID =   1738
            Width           =   1575
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "and"
            Height          =   255
            Left            =   480
            TabIndex        =   51
            Top             =   1560
            WhatsThisHelpID =   1739
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Special Chars"
         Height          =   2415
         Left            =   0
         TabIndex        =   49
         Top             =   240
         WhatsThisHelpID =   1725
         Width           =   1935
         Begin VB.OptionButton optAlpha 
            Caption         =   "Letter"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            WhatsThisHelpID =   1726
            Width           =   1215
         End
         Begin VB.OptionButton optDigit 
            Caption         =   "Digit"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   480
            WhatsThisHelpID =   1727
            Width           =   1215
         End
         Begin VB.OptionButton optAlphaNum 
            Caption         =   "Alphanumeric"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   720
            WhatsThisHelpID =   1728
            Width           =   1455
         End
         Begin VB.OptionButton optWhitespace 
            Caption         =   "Whitespace"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   960
            WhatsThisHelpID =   1729
            Width           =   1455
         End
         Begin VB.OptionButton optCharSet 
            Caption         =   "Set"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1440
            WhatsThisHelpID =   1731
            Width           =   1455
         End
         Begin VB.CheckBox chkCharSetNot 
            Caption         =   "NOT in set"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   2040
            WhatsThisHelpID =   1732
            Width           =   1455
         End
         Begin VB.OptionButton optAnyChar 
            Caption         =   "Any character"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1200
            WhatsThisHelpID =   1730
            Width           =   1455
         End
         Begin MSForms.TextBox txtCharSet 
            Height          =   375
            Left            =   480
            TabIndex        =   27
            Top             =   1680
            Width           =   1335
            VariousPropertyBits=   746604571
            Size            =   "2355;661"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   34
         End
      End
      Begin MSForms.TextBox txtChars2Insert 
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   4920
         Width           =   1695
         VariousPropertyBits=   746604571
         Size            =   "2990;661"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
   End
   Begin VB.CommandButton cmdHelpPat 
      Caption         =   "Help with patterns"
      Height          =   495
      Left            =   10440
      TabIndex        =   18
      Top             =   3840
      WhatsThisHelpID =   1723
      Width           =   1215
   End
   Begin VB.CheckBox chkIgnoreCase 
      Caption         =   "Use case-insensitive matching"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      WhatsThisHelpID =   1714
      Width           =   6615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   12480
      TabIndex        =   19
      Top             =   6120
      WhatsThisHelpID =   1712
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Payee Mapping Rule"
      Height          =   3615
      Left            =   120
      TabIndex        =   41
      Top             =   3360
      WhatsThisHelpID =   1706
      Width           =   8655
      Begin VB.CheckBox chkNoReturn 
         Caption         =   "Continue to next rule"
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   1680
         WhatsThisHelpID =   1749
         Width           =   5295
      End
      Begin VB.TextBox txtSIC 
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1320
         Width           =   5295
      End
      Begin VB.ComboBox cbField 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cbMode 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test"
         Height          =   375
         Left            =   7440
         TabIndex        =   13
         Top             =   2760
         WhatsThisHelpID =   1715
         Width           =   1095
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   7440
         TabIndex        =   12
         Top             =   840
         WhatsThisHelpID =   1711
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   7440
         TabIndex        =   11
         Top             =   360
         WhatsThisHelpID =   1710
         Width           =   1095
      End
      Begin MSForms.ComboBox cbCategory 
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   960
         Width           =   5295
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "9340;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTestRepFull 
         Height          =   375
         Left            =   2040
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   3120
         Width           =   5295
         VariousPropertyBits=   746604571
         BackColor       =   -2147483633
         Size            =   "9340;661"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTestRep 
         Height          =   375
         Left            =   2040
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2760
         Width           =   5295
         VariousPropertyBits=   746604571
         BackColor       =   -2147483633
         Size            =   "9340;661"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTestMemo 
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   2400
         Width           =   5295
         VariousPropertyBits=   746604571
         Size            =   "9340;661"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTestPayee 
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   2040
         Width           =   5295
         VariousPropertyBits=   746604571
         Size            =   "9340;661"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPattern 
         Height          =   375
         Left            =   4200
         TabIndex        =   4
         Top             =   240
         Width           =   3135
         VariousPropertyBits=   746604571
         Size            =   "5530;661"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtReplacement 
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   600
         Width           =   5295
         VariousPropertyBits=   746604571
         Size            =   "9340;661"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblSIC 
         Alignment       =   1  'Right Justify
         Caption         =   "SIC:"
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Top             =   1320
         WhatsThisHelpID =   1748
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Test memo:"
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   2400
         WhatsThisHelpID =   1747
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Result of all rules:"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   3120
         WhatsThisHelpID =   1718
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Test payee:"
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   2040
         WhatsThisHelpID =   1716
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Result of this rule:"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   2760
         WhatsThisHelpID =   1717
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Set the category to:"
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   960
         WhatsThisHelpID =   1709
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Set the payee to:"
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   600
         WhatsThisHelpID =   1708
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "If payee:"
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   240
         WhatsThisHelpID =   1707
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   10440
      TabIndex        =   16
      Top             =   2040
      WhatsThisHelpID =   1705
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   10440
      TabIndex        =   17
      Top             =   2520
      WhatsThisHelpID =   1704
      Width           =   1095
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Move Down"
      Height          =   375
      Left            =   10440
      TabIndex        =   15
      Top             =   600
      WhatsThisHelpID =   1703
      Width           =   1095
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Move Up"
      Height          =   375
      Left            =   10440
      TabIndex        =   14
      Top             =   120
      WhatsThisHelpID =   1702
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvPayeeMap 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      WhatsThisHelpID =   1701
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pattern"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Field"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Replacement"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Category"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "SIC"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Cont?"
         Object.Width           =   1111
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Save && Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   11040
      TabIndex        =   20
      Top             =   6120
      WhatsThisHelpID =   1713
      Width           =   1095
   End
End
Attribute VB_Name = "frmPayee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : frmPayee
'    Project    : MT2OFX
'
'    Description:
'
'    Modified   : $Author: Colin $ $Date: 27/03/11 23:33 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/frmPayee.frm 20    27/03/11 23:33 Colin $"
' $History: frmPayee.frm $
' 
' *****************  Version 20  *****************
' User: Colin        Date: 27/03/11   Time: 23:33
' Updated in $/MT2OFX
' multiple account support
'
' *****************  Version 19  *****************
' User: Colin        Date: 15/11/10   Time: 0:01
' Updated in $/MT2OFX
'
' *****************  Version 18  *****************
' User: Colin        Date: 30/08/09   Time: 13:19
' Updated in $/MT2OFX
'
' *****************  Version 16  *****************
' User: Colin        Date: 20/04/08   Time: 10:05
' Updated in $/MT2OFX
' For 3.5 beta 1
'
' *****************  Version 15  *****************
' User: Colin        Date: 7/12/06    Time: 14:59
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 14  *****************
' User: Colin        Date: 25/04/06   Time: 21:44
' Updated in $/MT2OFX
'
' *****************  Version 11  *****************
' User: Colin        Date: 2/11/05    Time: 23:03
' Updated in $/MT2OFX
' V3.4 beta 1
'
' *****************  Version 10  *****************
' User: Colin        Date: 8/05/05    Time: 12:43
' Updated in $/MT2OFX
' V3.3.8
'
' *****************  Version 9  *****************
' User: Colin        Date: 18/03/05   Time: 21:57
' Updated in $/MT2OFX
'
' *****************  Version 8  *****************
' User: Colin        Date: 6/03/05    Time: 23:41
' Updated in $/MT2OFX
'</CSCC>

#Const nohelp = False

Dim oPayees As PayeeMap
Dim bEditInProgress As Boolean
Dim mbIgnoreListClick As Boolean
Dim bInsertInProgress As Boolean

Private oRegExp As New RegExp
Private sRepl As String
Private oTempPayeeMap As New PayeeMap
Private oTempPayeeMapFull As New PayeeMap

Private gsStartsWith As String
Private gsContains As String
Private gsMatches As String
Private gsFieldMemo As String
Private gsFieldPayee As String

Private Const csNoReturnChar As String = "*"

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cbCategory_Change
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12/07/2004-21:54:08
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cbCategory_Change()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cbCategory_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12/07/2004-22:40:31
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cbCategory_Click()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cbField_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       14/02/2005-14:05:06
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cbField_Click()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cbMode_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       28-Jan-2005-00:25:23
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cbMode_Click()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       chkCapture_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       22/08/2004-14:46:18
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub chkCapture_Click()
    UpdateInsertText
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       chkCharSetNot_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-21:59:49
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub chkCharSetNot_Click()
    UpdateInsertText
End Sub

Private Sub chkNoReturn_Click()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdAdd_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12/07/2004-00:07:40
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdAdd_Click()
    Dim iAfter As Long
    bInsertInProgress = True
    If lvPayeeMap.SelectedItem Is Nothing Then
        iAfter = lvPayeeMap.ListItems.Count + 1
    Else
        iAfter = lvPayeeMap.SelectedItem.Index
    End If
    With Me.lvPayeeMap.ListItems.Add(iAfter, , "")
        .SubItems(1) = ""
        .SubItems(2) = ""
        .SubItems(3) = ""
        .SubItems(4) = ""
        .SubItems(5) = ""
        .Selected = True
    End With
    txtPattern = ""
    txtReplacement = ""
    cbCategory = ""
    UpdateUI
    txtPattern.SetFocus
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdCancel_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       11/07/2004-23:47:37
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
' Procedure  :       cmdClose_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       05/07/2004-20:27:55
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdClose_Click()
    LoadPayeeMapFromWindow oPayees
    If oPayees.SavePayeeMap(Cfg.PayeeMapFile) Then
        Cfg.PayeeMapIgnoreCase = (chkIgnoreCase.Value = vbChecked)
        Cfg.TestPayee = txtTestPayee
        Cfg.TestMemo = txtTestMemo
        Cfg.Save
        Unload Me
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdDown_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12/07/2004-00:21:00
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdDown_Click()
    Dim iTmp As Long
    Dim i As Long
    Dim sPattern As String, sReplacement As String, sCategory As String
    Dim sField As String
    Dim sSIC As String
    Dim sNoRet As String
    With lvPayeeMap
        With .SelectedItem
            iTmp = .Index
            sPattern = .Text
            sField = .SubItems(1)
            sReplacement = .SubItems(2)
            sCategory = .SubItems(3)
            sSIC = .SubItems(4)
            sNoRet = .SubItems(5)
        End With
        With .ListItems
            .Remove iTmp  ' remove selected item
                ' now insert before previous item
            With .Add(iTmp + 1, , sPattern)
                .SubItems(1) = sField
                .SubItems(2) = sReplacement
                .SubItems(3) = sCategory
                .SubItems(4) = sSIC
                .SubItems(5) = sNoRet
            End With
        End With
'        For i = iTmp To .ListItems.Count
'            .ListItems(i).Text = CStr(i)
'        Next
        Set .SelectedItem = .ListItems(iTmp + 1)
    End With
    cmdClose.Enabled = True
    UpdateUI
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdHelpPat_Click
' Description:       show help for patterns
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/08/2004-17:09:59
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdHelpPat_Click()
    ShowHelpTopic HH_Payee_Mapping_Window
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdInsertChars_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-22:42:07
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdInsertChars_Click()
    txtPattern.SelText = txtChars2Insert
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdRemove_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12/07/2004-00:07:54
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdRemove_Click()
    With lvPayeeMap
        .ListItems.Remove .SelectedItem.Index
    End With
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdReset_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       11/07/2004-23:47:28
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdReset_Click()
    Dim iItemIndex As Long
    DebugLog "In " & Me.Name & ".cmdReset_Click", logDEBUG
    If lvPayeeMap.SelectedItem Is Nothing Then
        txtPattern = ""
        txtReplacement = ""
        cbCategory = ""
        cbField.ListIndex = -1
    Else
        If bInsertInProgress Then   ' we are "inserting" a new entry
                                    ' so we can remove the empty one
            iItemIndex = lvPayeeMap.SelectedItem.Index
            lvPayeeMap.ListItems.Remove iItemIndex
            iItemIndex = iItemIndex - 1
            If iItemIndex < lvPayeeMap.ListItems.Count Then
                Set lvPayeeMap.SelectedItem = Nothing
            Else
                lvPayeeMap.ListItems(iItemIndex).Selected = True
            End If
        Else
            With lvPayeeMap.SelectedItem
                PatternFromList .Text
                cbField.Text = .SubItems(1)
                txtReplacement = .SubItems(2)
                cbCategory = .SubItems(3)
                txtSIC = .SubItems(4)
                chkNoReturn.Value = IIf(.SubItems(5) = "", vbUnchecked, vbChecked)
            End With
        End If
    End If
    bEditInProgress = False
    bInsertInProgress = False
    cmdReset.Enabled = False
    cmdUpdate.Enabled = False
    UpdateUI
    DebugLog "Leaving " & Me.Name & ".cmdReset_Click", logDEBUG
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdTest_Click
' Description:       click on test button - see what the matching returns
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/08/2004-14:26:49
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdTest_Click()
    TestMatch
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdUp_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12/07/2004-00:08:11
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdUp_Click()
    Dim iTmp As Long
    Dim i As Long
    Dim sPattern As String, sReplacement As String, sCategory As String
    Dim sField As String, sSIC As String, sNoRet As String
    With lvPayeeMap
        With .SelectedItem
            iTmp = .Index
            sPattern = .Text
            sField = .SubItems(1)
            sReplacement = .SubItems(2)
            sCategory = .SubItems(3)
            sSIC = .SubItems(4)
            sNoRet = .SubItems(5)
        End With
        With .ListItems
            .Remove iTmp  ' remove selected item
                ' now insert before previous item
            With .Add(iTmp - 1, , sPattern)
                .SubItems(1) = sField
                .SubItems(2) = sReplacement
                .SubItems(3) = sCategory
                .SubItems(4) = sSIC
                .SubItems(5) = sNoRet
            End With
        End With
'        For i = iTmp To .ListItems.Count
'            .ListItems(i).Text = CStr(i)
'        Next
        Set .SelectedItem = .ListItems(iTmp - 1)
    End With
    cmdClose.Enabled = True
    UpdateUI
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdUpdate_Click
' Description:       [type_description_here]
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12/07/2004-00:08:08
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdUpdate_Click()
    Dim iSIC As Long
    If Not IsValidPattern() Then
        If MyMsgBox(LoadResStringL(1719), vbOKCancel, LoadResStringL(1720)) <> vbOK Then
            txtPattern.SetFocus
            Exit Sub
        End If
    End If
    If lvPayeeMap.SelectedItem Is Nothing Then  ' this should not happen!
        Debug.Assert False
        Exit Sub
    End If
    ' accept pattern
    With lvPayeeMap.SelectedItem
        Select Case cbMode
        Case gsStartsWith
            .Text = "^" & txtPattern
        Case gsContains
            .Text = txtPattern
        Case Else   ' matches
            .Text = "^" & txtPattern & "$"
        End Select
        .SubItems(1) = cbField
        .SubItems(2) = txtReplacement
        .SubItems(3) = cbCategory
        iSIC = Val(txtSIC)
        If iSIC > 0 Then
            .SubItems(4) = CStr(iSIC)
            txtSIC = CStr(iSIC)
        Else
            .SubItems(4) = ""
        End If
        .SubItems(5) = IIf(chkNoReturn.Value = vbChecked, csNoReturnChar, "")
    End With
    AddTextToCombo cbCategory, cbCategory.Text
    cmdReset.Enabled = False
    cmdUpdate.Enabled = False
    bEditInProgress = False
    bInsertInProgress = False
    UpdateUI
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Form_KeyDown
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/08/2004-21:32:05
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        ShowHelpTopic HH_Payee_Mapping_Window
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
' Date-Time  :       07/07/2004-00:17:22
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()
    Dim sTmp As String
    DebugLog "In " & Me.Name & ".Load", logDEBUG
    LocaliseForm Me
    DebugLog "Done localisation", logDEBUG
    Me.txtPattern = ""
    Me.txtReplacement = ""
    Me.txtTestPayee = Cfg.TestPayee
    Me.txtTestMemo = Cfg.TestMemo
    Me.txtTestRep = ""
    Me.txtTestRepFull = ""
    Me.txtCharSet = ""
    Me.txtNumberFrom = ""
    Me.txtNumberTo = ""
    Me.txtSIC = ""
    Me.lvPayeeMap.ListItems.Clear
    Me.cbMode.Clear
    gsStartsWith = LoadResStringL(1742)
    gsContains = LoadResStringL(1743)
    gsMatches = LoadResStringL(1744)
    gsFieldPayee = LoadResStringL(1745)
    gsFieldMemo = LoadResStringL(1746)
    DebugLog "Populating cbMode", logDEBUG
    With Me.cbMode
        .Clear
        .AddItem gsStartsWith
        .AddItem gsContains
        .AddItem gsMatches
    End With
    DebugLog "Populating cbField", logDEBUG
    With Me.cbField
        .Clear
        .AddItem gsFieldMemo
        .AddItem gsFieldPayee
    End With
    Set Me.lvPayeeMap.SelectedItem = Nothing
    Debug.Assert (lvPayeeMap.SelectedItem Is Nothing)
    Me.cbCategory.Clear
    Set oPayees = New PayeeMap
    DebugLog "Loading Payee Map from " & Cfg.PayeeMapFile
    oPayees.LoadPayeeMap Cfg.PayeeMapFile
    Me.chkIgnoreCase.Value = IIf(Cfg.PayeeMapIgnoreCase, vbChecked, vbUnchecked)
    Dim oPayee As PayeeMapItem
    DebugLog "Loading Payee Map control"
    For Each oPayee In oPayees
        With lvPayeeMap.ListItems.Add(, , oPayee.Pattern)
            Select Case oPayee.MatchField
            Case mfMemo
                sTmp = gsFieldMemo
            Case mfPayee
                sTmp = gsFieldPayee
            Case Else
                Debug.Assert False
                sTmp = gsFieldPayee
            End Select
            .SubItems(1) = sTmp
            .SubItems(2) = oPayee.Replacement
            .SubItems(3) = oPayee.Category
            If oPayee.SIC > 0 Then
                .SubItems(4) = CStr(oPayee.SIC)
            Else
                .SubItems(4) = ""
            End If
            .SubItems(5) = IIf(oPayee.NoReturn, csNoReturnChar, "")
            .Selected = False
        End With
        AddTextToCombo cbCategory, oPayee.Category
    Next
    DebugLog "Done loading listview", logDEBUG
    txtPattern = ""
    Set Me.lvPayeeMap.SelectedItem = Nothing
    txtReplacement = ""
    cbCategory.Text = ""
    bEditInProgress = False
    Call cmdReset_Click
    UpdateUI
    DebugLog "Leaving " & Me.Name & ".Load", logDEBUG
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       UpdateUI
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       07/07/2004-00:18:00
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub UpdateUI()
    Dim bSelection As Boolean
    DebugLog "In " & Me.Name & ".UpdateUI", logDEBUG
    If bEditInProgress Then
        cmdUp.Enabled = False
        cmdDown.Enabled = False
        cmdRemove.Enabled = False
        cmdAdd.Enabled = False
        cmdClose.Enabled = False
        cbMode.Enabled = True
        txtPattern.Enabled = True
        txtReplacement.Enabled = True
        cbCategory.Enabled = True
        cbField.Enabled = True
        lvPayeeMap.Enabled = False
    Else
        bSelection = Not (lvPayeeMap.SelectedItem Is Nothing)
        If bSelection Then
            cmdUp.Enabled = lvPayeeMap.SelectedItem.Index <> 1
            cmdDown.Enabled = lvPayeeMap.SelectedItem.Index <> lvPayeeMap.ListItems.Count
            cbMode.Enabled = True
            txtPattern.Enabled = True
            txtReplacement.Enabled = True
            cbCategory.Enabled = True
            cbField.Enabled = True
        Else
            cmdUp.Enabled = False
            cmdDown.Enabled = False
            cbMode.Enabled = False
            txtPattern.Enabled = False
            txtReplacement.Enabled = False
            cbCategory.Enabled = False
            cbField.Enabled = False
        End If
        lvPayeeMap.Enabled = True
        cmdRemove.Enabled = bSelection
        cmdAdd.Enabled = True
        cmdClose.Enabled = True
    End If
    cmdInsertChars.Enabled = txtPattern.Enabled And (Len(txtChars2Insert) > 0)
    cmdTest.Enabled = (Len(txtPattern) > 0) And ((Len(txtTestPayee) > 0) Or (Len(txtTestMemo) > 0))
    DebugLog "Leaving " & Me.Name & ".UpdateUI", logDEBUG
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       StartEdit
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12/07/2004-22:50:57
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Sub StartEdit()
    cmdReset.Enabled = True
    cmdUpdate.Enabled = (Len(txtPattern) > 0)
    bEditInProgress = True
    UpdateUI
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Form_Resize
' Description:      temp code to test special char handling!!!
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-17:45:56
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
#If nohelp Then
Private Sub Form_Resize()
    If Not InIDE(Me.hWnd) Then
        If Me.Width > 9165 Then
            Me.Width = 9165
        End If
    End If
End Sub
#End If

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       lvPayeeMap_Click
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       07/07/2004-00:22:10
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub lvPayeeMap_Click()
    UpdateUI
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       lvPayeeMap_ItemClick
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       07/07/2004-00:22:31
'
' Parameters :       Item (MSComctlLib.ListItem)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub lvPayeeMap_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    mbIgnoreListClick = True
    PatternFromList Item.Text
    cbField.Text = Item.SubItems(1)
    txtReplacement = Item.SubItems(2)
    cbCategory.Text = Item.SubItems(3)
    txtSIC = Item.SubItems(4)
    chkNoReturn.Value = IIf(Item.SubItems(5) <> "", vbChecked, vbUnchecked)
    txtTestRep = ""
    cmdReset.Enabled = False
    cmdUpdate.Enabled = False
    bEditInProgress = False
    UpdateUI
    mbIgnoreListClick = True
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       optAlpha_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-21:59:12
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub optAlpha_Click()
    UpdateInsertText
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       optAlphaNum_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-22:00:12
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub optAlphaNum_Click()
    UpdateInsertText
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       optAnyChar_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-22:00:29
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub optAnyChar_Click()
    UpdateInsertText
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       optBetween_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-22:00:43
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub optBetween_Click()
    UpdateInsertText
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       optCharSet_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-22:01:44
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub optCharSet_Click()
    UpdateInsertText
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       optDigit_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-22:02:06
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub optDigit_Click()
    UpdateInsertText
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       optOne_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-22:02:18
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub optOne_Click()
    UpdateInsertText
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       optOneOrMore_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-22:02:32
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub optOneOrMore_Click()
    UpdateInsertText
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       optWhitespace_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-22:02:46
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub optWhitespace_Click()
    UpdateInsertText
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       optZeroOrMore_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-22:02:58
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub optZeroOrMore_Click()
    UpdateInsertText
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       optZeroOrOne_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-22:03:17
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub optZeroOrOne_Click()
    UpdateInsertText
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtChars2Insert_Change
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       28-Jan-2005-12:12:39
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtChars2Insert_Change()
    UpdateUI
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtCharSet_Change
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-22:03:34
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtCharSet_Change()
    UpdateInsertText
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtNumberFrom_Change
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-22:03:49
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtNumberFrom_Change()
    UpdateInsertText
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtNumberFrom_KeyPress
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-22:36:18
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtNumberFrom_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtNumberTo_Change
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-22:04:01
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtNumberTo_Change()
    UpdateInsertText
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtNumberTo_KeyPress
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-22:36:53
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtNumberTo_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtPattern_Change
' Description:       [type_description_here]
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       11/07/2004-23:35:00
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtPattern_Change()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtReplacement_Change
' Description:       [type_description_here]
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       11/07/2004-23:35:22
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtReplacement_Change()
    StartEdit
End Sub


'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       IsValidPattern
' Description:       validate the pattern
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       19/08/2004-23:50:20
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Function IsValidPattern() As Boolean
    On Error GoTo badpat
    oRegExp.Pattern = txtPattern
    oRegExp.Global = False
    oRegExp.MultiLine = False
    oRegExp.IgnoreCase = (Me.chkIgnoreCase.Value = vbChecked)
    Dim ms As MatchCollection
    Set ms = oRegExp.Execute("bla")
    IsValidPattern = True
goback:
    Exit Function
badpat:
    IsValidPattern = False
    Resume goback
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       TestMatch
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/08/2004-00:16:13
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub TestMatch()
    Dim sTmp As String
    If IsValidPattern() Then
        If DoReplacement() Then
            txtTestRep = sRepl
        Else
            txtTestRep = LoadResStringL(1721)
        End If
    Else
        txtTestRep = LoadResStringL(1722)
    End If
    Dim oItem As PayeeMapItem
    LoadPayeeMapFromWindow oTempPayeeMapFull
    oTempPayeeMapFull.IgnoreCase = (chkIgnoreCase.Value = vbChecked)
' 20050214 CS: added new first param (memo) to MapSearch to support matching on other fields
    Set oItem = oTempPayeeMapFull.MapSearch(txtTestMemo, txtTestPayee, sTmp)
    If oItem Is Nothing Then
        txtTestRepFull = LoadResStringL(1721)
    Else
        txtTestRepFull = sTmp
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoReplacement
' Description:       perform test replacement to see what comes out
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/08/2004-13:54:51
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Function DoReplacement() As Boolean
    Dim sTmp As String
    Dim oItem As New PayeeMapItem
    
    Select Case cbMode
    Case gsStartsWith
        sTmp = "^" & txtPattern
    Case gsContains
        sTmp = txtPattern
    Case Else   ' matches
        sTmp = "^" & txtPattern & "$"
    End Select
    oItem.Pattern = sTmp
    
    Select Case cbField
    Case gsFieldMemo
        oItem.MatchField = mfMemo
    Case gsFieldPayee
        oItem.MatchField = mfPayee
    Case Else
        oItem.MatchField = mfPayee
        Debug.Assert False
    End Select
    oItem.Replacement = txtReplacement
    oTempPayeeMap.Clear
    oTempPayeeMap.Add oItem
    oTempPayeeMap.IgnoreCase = (chkIgnoreCase.Value = vbChecked)
    Dim oItemOut As PayeeMapItem
' 20050214 CS: added new first param (memo) to MapSearch to support matching on other fields
    Set oItemOut = oTempPayeeMap.MapSearch(txtTestMemo, txtTestPayee, sRepl)
    If oItemOut Is Nothing Then
        DoReplacement = False
    Else
        DoReplacement = True
    End If
End Function
'enter pattern element

Private Sub txtSIC_Change()
    StartEdit
End Sub

Private Sub txtSIC_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    'Allows only numeric keys and (backspace) keys
    'and removes any others by setting it to null = chr(0)
    Case 8, 48 To 57, 127
    Case Else: KeyAscii = 0
    End Select
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtTestMemo_Change
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       14/02/2005-15:15:06
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtTestMemo_Change()
    txtTestRep = ""
    txtTestRepFull = ""
    UpdateUI
End Sub

'alpha
'digit
'Space
'any char
'char set
'not char set

'one
'zero or more
'one or more
'at least X
'at most Y

'capture expression

'start
'End

'literal text (escape things!)


Private Sub txtTestPayee_Change()
    txtTestRep = ""
    txtTestRepFull = ""
    UpdateUI
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       LoadPayeeMapFromWindow
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/08/2004-16:33:53
'
' Parameters :       pm (PayeeMap)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub LoadPayeeMapFromWindow(pm As PayeeMap)
    pm.Clear
    Dim i As Long
    Dim oMap As PayeeMapItem
    Dim oItem As ListItem
    For Each oItem In lvPayeeMap.ListItems
        Set oMap = New PayeeMapItem
        With oItem
            oMap.Pattern = .Text
            Select Case .SubItems(1)
            Case gsFieldMemo
                oMap.MatchField = mfMemo
            Case gsFieldPayee
                oMap.MatchField = mfPayee
            End Select
            oMap.Replacement = .SubItems(2)
            oMap.Category = .SubItems(3)
            oMap.SIC = Val(.SubItems(4))
            oMap.NoReturn = (.SubItems(5) <> "")
        End With
        pm.Add oMap
    Next
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       UpdateInsertText
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/08/2004-21:59:38
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub UpdateInsertText()
    Dim sTmp As String
    Dim sTmp2 As String
    Dim bTmp As Boolean
    bTmp = False
    If optAlpha Then
        sTmp = "[A-Za-z]"
    ElseIf optDigit Then
        sTmp = "\d"
    ElseIf optAnyChar Then
        sTmp = "."
    ElseIf optCharSet Then
        sTmp2 = txtCharSet
        sTmp2 = Replace(sTmp2, "[", "\[")
        sTmp2 = Replace(sTmp2, "]", "\]")
        If chkCharSetNot.Value = vbChecked Then
            sTmp = "[^" & sTmp2 & "]"
        Else
            sTmp = "[" & sTmp2 & "]"
        End If
        bTmp = True
    ElseIf optAlphaNum Then
        sTmp = "[A-Za-z\d]"
    ElseIf optWhitespace Then
        sTmp = "\s"
    End If
    txtCharSet.Enabled = bTmp
    chkCharSetNot.Enabled = bTmp
    
    bTmp = False
    If optOne Then
        sTmp = sTmp
    ElseIf optOneOrMore Then
        sTmp = sTmp & "+"
    ElseIf optZeroOrMore Then
        sTmp = sTmp & "*"
    ElseIf optZeroOrOne Then
        sTmp = sTmp & "?"
    ElseIf optBetween Then
        bTmp = True
        If txtNumberFrom = txtNumberTo Then
            sTmp = sTmp & "{" & txtNumberFrom & "}"
        Else
            sTmp = sTmp & "{" & txtNumberFrom & "," & txtNumberTo & "}"
        End If
    End If
    If chkCapture.Value = vbChecked Then
        sTmp = "(" & sTmp & ")"
    End If
    txtNumberFrom.Enabled = bTmp
    txtNumberTo.Enabled = bTmp
    
    txtChars2Insert = sTmp
'    UpdateUI
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       PatternFromList
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       28-Jan-2005-12:15:25
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub PatternFromList(sPat As String)
    If Left$(sPat, 1) = "^" And Right$(sPat, 1) = "$" Then
        txtPattern = Mid$(sPat, 2, Len(sPat) - 2)
        cbMode.ListIndex = FindInCombo(cbMode, gsMatches)
'        cbMode.ListIndex = 2    ' matches
    ElseIf Left$(sPat, 1) = "^" Then
        txtPattern = Mid$(sPat, 2)
        cbMode.ListIndex = FindInCombo(cbMode, gsStartsWith)
'        cbMode.ListIndex = 0    ' starts with
    Else
        txtPattern = sPat
        cbMode.ListIndex = FindInCombo(cbMode, gsContains)
'        cbMode.ListIndex = 1    ' contains
    End If
End Sub

