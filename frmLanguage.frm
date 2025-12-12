VERSION 5.00
Begin VB.Form frmLanguage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose Language"
   ClientHeight    =   1065
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbLanguages 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Please choose a language:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : frmLanguage
'    Project    : MT2OFX
'
'    Description: Language Selection Form
'
'    Modified   : $Author: Colin $ $Date: 14/11/10 23:54 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/frmLanguage.frm 8     14/11/10 23:54 Colin $"
' $History: frmLanguage.frm $
' 
' *****************  Version 8  *****************
' User: Colin        Date: 14/11/10   Time: 23:54
' Updated in $/MT2OFX
'
' *****************  Version 6  *****************
' User: Colin        Date: 7/12/06    Time: 13:19
' Updated in $/MT2OFX
' Path Management
'
' *****************  Version 3  *****************
' User: Colin        Date: 18/03/05   Time: 21:57
' Updated in $/MT2OFX
'
' *****************  Version 2  *****************
' User: Colin        Date: 6/03/05    Time: 0:35
' Updated in $/MT2OFX
'</CSCC>

Private aFileList(1 To 100) As String
Private iFileCount As Integer
Private Const gsEnglishKey As String = "#ENG"
Private Const gsDutchKey As String = "#DUT"
#Const Dutch_Builtin = False

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cbLanguages_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       04/03/2005-23:11:59
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cbLanguages_Click()
    UpdateUI
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdCancel_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/02/2005-22:23:08
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
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/02/2005-22:24:07
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdOK_Click()
    Dim iSel As Long
    iSel = cbLanguages.ItemData(cbLanguages.ListIndex)
    Dim sFile As String
    Dim bRet As Boolean
    Dim lLocale As Long
    If iSel = 0 Then
        sFile = ""
        lLocale = &H809
#If Dutch_Builtin Then
    ElseIf iSel = -1 Then
        lLocale = &H413
        sFile = ""
#End If
    Else
        sFile = aFileList(iSel)
        lLocale = 0
    End If
    bRet = SetLanguageFile(sFile)
    If bRet Then
        If lLocale > 0 Then
            SetProgLocale lLocale
        End If
        Cfg.ResFile = sFile
        Cfg.Save
    End If
    Unload Me
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Form_Load
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/02/2005-22:29:56
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()
    Dim sDir As String
    Dim sFile As String
    Dim sLang As String
    Dim sLCID As String
    Dim lLocale As Long
    Dim lLCID As String
    Dim i As Integer
#If Dutch_Builtin Then
    iFileCount = 2
    aFileList(1) = gsEnglishKey
    aFileList(2) = gsDutchKey
#Else
    iFileCount = 1
    aFileList(1) = gsEnglishKey
#End If
    sDir = Cfg.ResourcePath & "\"
    sFile = Dir(sDir & "*.*")
    Do While sFile <> ""
        iFileCount = iFileCount + 1
        aFileList(iFileCount) = sFile
        sFile = Dir
    Loop
    lLocale = GetProgLocale()
    cbLanguages.ListIndex = -1
    cbLanguages.Clear
    cbLanguages.AddItem "English"
    cbLanguages.ItemData(cbLanguages.newIndex) = 0
    If PRIMARYLANGID(lLocale) = LANG_ENGLISH Then
        cbLanguages.ListIndex = cbLanguages.newIndex
    End If
#If Dutch_Builtin Then
    cbLanguages.AddItem "Nederlands"
    cbLanguages.ItemData(cbLanguages.newIndex) = -1
    If PRIMARYLANGID(lLocale) = LANG_DUTCH Then
        cbLanguages.ListIndex = cbLanguages.newIndex
    End If
#End If
#If Dutch_Builtin Then
    For i = 3 To iFileCount
#Else
    For i = 2 To iFileCount
#End If
        sFile = aFileList(i)
        sLang = ReadIniString(sDir & sFile, "MT2OFX Language File", "Language")
        If sLang <> "" Then
            cbLanguages.AddItem sLang
            cbLanguages.ItemData(cbLanguages.newIndex) = i
            sLCID = ReadIniString(sDir & sFile, "MT2OFX Language File", "LCID")
            If IsNumeric(sLCID) Then
                lLCID = CLng(sLCID)
                If lLCID = lLocale Then
                    cbLanguages.ListIndex = cbLanguages.newIndex
                End If
            End If
        End If
    Next
    cmdOK.Enabled = False
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       UpdateUI
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       04/03/2005-23:12:49
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub UpdateUI()
    cmdOK.Enabled = (cbLanguages.ListIndex <> -1)
End Sub

