VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmScriptCfg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Script Properties"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7185
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   375
      Left            =   6360
      TabIndex        =   16
      Top             =   1440
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Property Value"
      Height          =   1695
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      WhatsThisHelpID =   1805
      Width           =   6975
      Begin VB.ComboBox cbChoiceValue 
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtFloatValue 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Text            =   "number"
         Top             =   1080
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtDateValue 
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   173867009
         CurrentDate     =   38311
      End
      Begin VB.CheckBox chkBoolValue 
         Caption         =   "Check1"
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H8000000F&
         Height          =   735
         Left            =   1320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "frmScriptCfg.frx":0000
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox txtNumValue 
         Height          =   375
         Left            =   5040
         TabIndex        =   4
         Text            =   "number"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtTextValue 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Text            =   "string"
         Top             =   1200
         Width           =   4455
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   5760
         TabIndex        =   11
         Top             =   1080
         WhatsThisHelpID =   1809
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   5760
         TabIndex        =   10
         Top             =   600
         WhatsThisHelpID =   1808
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         WhatsThisHelpID =   1806
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Value:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         WhatsThisHelpID =   1807
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lvProps 
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   480
      WhatsThisHelpID =   1804
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Setting"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   6174
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   600
      WhatsThisHelpID =   1803
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   120
      WhatsThisHelpID =   1802
      Width           =   1095
   End
   Begin VB.Label lblScript 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   120
      WhatsThisHelpID =   1801
      Width           =   5775
   End
End
Attribute VB_Name = "frmScriptCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : frmScriptCfg
'    Project    : MT2OFX
'
'    Description:
'
'    Modified   : $Author: Colin $ $Date: 27/03/11 23:33 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/frmScriptCfg.frm 22    27/03/11 23:33 Colin $"
' $History: frmScriptCfg.frm $
' 
' *****************  Version 22  *****************
' User: Colin        Date: 27/03/11   Time: 23:33
' Updated in $/MT2OFX
' multiple account support
'
' *****************  Version 21  *****************
' User: Colin        Date: 15/11/10   Time: 0:04
' Updated in $/MT2OFX
'
' *****************  Version 20  *****************
' User: Colin        Date: 28/10/09   Time: 22:52
' Updated in $/MT2OFX
'
' *****************  Version 19  *****************
' User: Colin        Date: 15/06/09   Time: 19:24
' Updated in $/MT2OFX
' For transfer to new laptop
'
' *****************  Version 18  *****************
' User: Colin        Date: 25/11/08   Time: 22:22
' Updated in $/MT2OFX
' moving vss server!
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
' User: Colin        Date: 6/05/05    Time: 23:15
' Updated in $/MT2OFX
' Changed control widths for German localisation
'
' *****************  Version 9  *****************
' User: Colin        Date: 6/03/05    Time: 23:41
' Updated in $/MT2OFX
'</CSCC>

Public PropSetName As String
Public Props As Variant
Public RetVal As Boolean
Public PropSet As ScriptPropertySet

Private sTrue As String
Private sFalse As String
Private bEditing As Boolean
Private iDecimalChar As Integer
Private aCurrencyList() As String

Private gsClose As String
Private gsCancel As String

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       StartEdit
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/11/2004-23:18:14
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub StartEdit()
    If Not bEditing Then
        bEditing = True
        cmdReset.Enabled = True
        cmdUpdate.Enabled = True
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GridToForm
' Description:       Fill controls from selected item in the grid
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/11/2004-00:05:56
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub GridToForm(Item As ListItem)
    Dim xProp As ScriptProperty
    Dim xChoice As Variant
    txtTextValue.Enabled = False
    txtTextValue.Visible = False
    txtNumValue.Enabled = False
    txtNumValue.Visible = False
    chkBoolValue.Enabled = False
    chkBoolValue.Visible = False
    dtDateValue.Enabled = False
    dtDateValue.Visible = False
    txtFloatValue.Enabled = False
    txtFloatValue.Visible = False
    cbChoiceValue.Enabled = False
    cbChoiceValue.Visible = False
    If Item Is Nothing Then
        txtDescription = ""
        txtTextValue = ""
        txtNumValue = ""
        txtFloatValue = ""
        cbChoiceValue.Clear
    Else
        Set xProp = PropSet(Item.Tag)
        txtDescription = xProp.Description
        Select Case xProp.DataType
        Case ptString
            txtTextValue.Enabled = True
            txtTextValue.Visible = True
            txtTextValue = Item.SubItems(1)
        Case ptInteger
            txtNumValue.Enabled = True
            txtNumValue.Visible = True
            txtNumValue = Item.SubItems(1)
        Case ptBoolean
            chkBoolValue.Enabled = True
            chkBoolValue.Visible = True
            chkBoolValue.Caption = xProp.Name
            chkBoolValue.Value = IIf(Item.SubItems(1) = sTrue, vbChecked, vbUnchecked)
        Case ptDate
            dtDateValue.Enabled = True
            dtDateValue.Visible = True
            If IsDate(Item.SubItems(1)) Then
                dtDateValue.Value = CDate(Item.SubItems(1))
            Else
                dtDateValue.Value = Now()
            End If
        Case ptFloat
            txtFloatValue.Enabled = True
            txtFloatValue.Visible = True
            txtFloatValue = Item.SubItems(1)
        Case ptChoice
            cbChoiceValue.Enabled = True
            cbChoiceValue.Visible = True
            cbChoiceValue.Clear
            For Each xChoice In xProp.Choices
                cbChoiceValue.AddItem CStr(xChoice)
            Next
            If Len(Trim(Item.SubItems(1))) > 0 Then
                cbChoiceValue.ListIndex = FindInCombo(cbChoiceValue, Item.SubItems(1))
            Else
                cbChoiceValue.ListIndex = -1
            End If
        Case ptCurrency
            cbChoiceValue.Enabled = True
            cbChoiceValue.Visible = True
            cbChoiceValue.Clear
            For Each xChoice In aCurrencyList
                If FindInCombo(cbChoiceValue, CStr(xChoice)) = -1 Then
                    cbChoiceValue.AddItem CStr(xChoice)
                End If
            Next
            If Len(Trim(Item.SubItems(1))) > 0 Then
                cbChoiceValue.ListIndex = FindInCombo(cbChoiceValue, Item.SubItems(1))
            Else
                cbChoiceValue.ListIndex = -1
            End If
        End Select
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       PropValString
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/11/2004-00:21:22
'
' Parameters :       xProp (ScriptProperty)
'--------------------------------------------------------------------------------
'</CSCM>
Private Function PropValString(xProp As ScriptProperty) As String
    Dim sTmp As String
    Select Case xProp.DataType
    Case ptString
        sTmp = xProp.Value
    Case ptInteger
        sTmp = CStr(xProp.Value)
    Case ptBoolean
        sTmp = IIf((xProp.Value = 0), sFalse, sTrue)
    Case ptDate
        sTmp = Format(xProp.Value, "Short Date", vbUseSystemDayOfWeek, vbUseSystem)
    Case ptFloat
        sTmp = CStr(xProp.Value)
    Case ptChoice
        sTmp = CStr(xProp.Value)
    Case ptCurrency
        sTmp = CStr(xProp.Value)
    End Select
    PropValString = sTmp
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       FormToGrid
' Description:       Update selected grid item from changed form members
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/11/2004-00:15:58
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub FormToGrid()
    Dim xProp As ScriptProperty
    Set xProp = PropSet(lvProps.SelectedItem.Tag)
    With lvProps.SelectedItem
        Select Case xProp.DataType
        Case ptString
            .SubItems(1) = txtTextValue
        Case ptInteger
            .SubItems(1) = txtNumValue
        Case ptBoolean
            .SubItems(1) = IIf(chkBoolValue.Value = vbChecked, sTrue, sFalse)
        Case ptDate
            .SubItems(1) = Format(dtDateValue, "Short Date", vbUseSystemDayOfWeek, vbUseSystem)
        Case ptFloat
            .SubItems(1) = txtFloatValue
        Case ptChoice
            .SubItems(1) = cbChoiceValue
        Case ptCurrency
            .SubItems(1) = Left$(cbChoiceValue, 3)
        End Select
'        .SubItems(1) = PropValString(xProp)
    End With
    cmdOK.Enabled = True
    cmdCancel.Caption = gsCancel
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cbChoiceValue_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/11/2004-16:54:51
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cbChoiceValue_Click()
    StartEdit
End Sub

Private Sub cbChoiceValue_Validate(Cancel As Boolean)
    Cancel = Not Validate(cbChoiceValue)
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       chkBoolValue_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/11/2004-23:17:21
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub chkBoolValue_Click()
    StartEdit
End Sub

Private Sub chkBoolValue_Validate(Cancel As Boolean)
    Cancel = Not Validate(chkBoolValue)
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdCancel_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       13/11/2004-00:25:40
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdCancel_Click()
    RetVal = False
    Me.Hide
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdOK_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       13/11/2004-00:26:23
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdOK_Click()
' get the values from the grid into the propertyset!
    Dim Item As ListItem
    Dim xProp As ScriptProperty
    Dim sTmp As String
    If Not InIDE(Me.hWnd) Then
        On Error Resume Next
    End If
    For Each Item In lvProps.ListItems
        sTmp = Item.SubItems(1)
        Set xProp = PropSet(Item.Tag)
        Select Case xProp.DataType
        Case ptString
            xProp.Value = sTmp
        Case ptInteger
            If Len(sTmp) = 0 Then
                xProp.Value = 0
            Else
                xProp.Value = CLng(sTmp)
            End If
        Case ptFloat
            If Len(sTmp) = 0 Then
                xProp.Value = 0#
            Else
                xProp.Value = CDbl(sTmp)
            End If
        Case ptBoolean
            xProp.Value = IIf(sTmp = sTrue, True, False)
        Case ptDate
            If IsDate(sTmp) Then
                xProp.Value = CDate(sTmp)
            Else
                xProp.Value = ""
            End If
        Case ptChoice
            xProp.Value = sTmp
        Case ptCurrency
            xProp.Value = Left$(sTmp, 3)
        End Select
    Next
    RetVal = True
    Me.Hide
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdReset_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/11/2004-23:21:38
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdReset_Click()
    GridToForm lvProps.SelectedItem
    bEditing = False
    cmdReset.Enabled = False
    cmdUpdate.Enabled = False
    lvProps.SetFocus
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdUpdate_Click
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/11/2004-23:21:52
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdUpdate_Click()
    FormToGrid
    bEditing = False
    cmdReset.Enabled = False
    cmdUpdate.Enabled = False
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       dtDateValue_Change
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/11/2004-23:19:48
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub dtDateValue_Change()
    StartEdit
End Sub

Private Sub dtDateValue_Validate(Cancel As Boolean)
    Cancel = Not Validate(dtDateValue)
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Form_Load
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       13/11/2004-00:23:52
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()
    LocaliseForm Me, 1800
    gsClose = LoadResStringL(1810)
    gsCancel = LoadResStringL(1803)
    sTrue = CStr(True)
    sFalse = CStr(False)
    iDecimalChar = Asc(SystemDecimalSeparator())
    aCurrencyList = GetCurrencyList
    cmdOK.Enabled = False
    cmdTest.Enabled = False
    cmdTest.Visible = False
    cmdReset.Enabled = False
    cmdUpdate.Enabled = False
    RepositionControls "txtTextValue", "txtNumValue", "chkBoolValue", "dtDateValue", "txtFloatValue", "cbChoiceValue"
' 20070306 CS: make sure & displays correctly and is not taken as an access key indicator
    lblScript.Caption = Replace(LoadResStringLEx(1801, PropSetName), "&", "&&")
    lvProps.ListItems.Clear
    txtDescription = ""
    If IsEmpty(Props) Then
        Exit Sub
    End If
    If PropSetName = "" Or (PropSet Is Nothing) Then
        Exit Sub
    End If
    If TypeName(Props) <> "Variant()" Then
        Exit Sub
    End If
    Dim xProp As ScriptProperty
    For Each xProp In PropSet
        With lvProps.ListItems.Add(, , xProp.Name)
            .ToolTipText = xProp.Description
            .Tag = xProp.Key
            .SubItems(1) = PropValString(xProp)
            .Selected = False
        End With
    Next
    Set lvProps.SelectedItem = Nothing
    Call lvProps_ItemClick(Nothing)
    cmdCancel.Caption = gsClose
End Sub

Private Sub lvProps_ItemClick(ByVal Item As MSComctlLib.ListItem)
    GridToForm Item
    bEditing = False
    cmdReset.Enabled = False
    cmdUpdate.Enabled = False
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtFloatValue_Change
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/11/2004-16:55:13
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtFloatValue_Change()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtFloatValue_KeyPress
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/11/2004-16:59:13
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtFloatValue_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    'Allows only numeric keys and (backspace) and (-) and (.) kesys
    'and removes any others by setting it to null = chr(0)
    Case 8, 45, 48 To 57, 127, iDecimalChar
    Case Else: KeyAscii = 0
    End Select
'Only allow the (-) sign as the first character
    If txtFloatValue.SelStart <> 0 And KeyAscii = 45 Then KeyAscii = 0
End Sub

Private Sub txtFloatValue_Validate(Cancel As Boolean)
    Cancel = Not Validate(txtFloatValue)
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtNumValue_Change
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/11/2004-23:19:05
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtNumValue_Change()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtNumValue_KeyPress
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       21/11/2004-00:01:13
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtNumValue_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    'Allows only numeric keys and (backspace) and (-) keys
    'and removes any others by setting it to null = chr(0)
    Case 8, 45, 48 To 57, 127
    Case Else: KeyAscii = 0
    End Select
'Only allow the (-) sign as the first character
    If txtNumValue.SelStart <> 0 And KeyAscii = 45 Then KeyAscii = 0
End Sub

Private Sub txtNumValue_Validate(Cancel As Boolean)
    Cancel = Not Validate(txtNumValue)
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtTextValue_Change
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/11/2004-23:19:24
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtTextValue_Change()
    StartEdit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       RepositionControls
' Description:       overlay all input controls onto the text control
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       15 Dec 2004-15:08:52
'
' Parameters :       aControls() (String)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub RepositionControls(sBase As String, ParamArray aControls())
    Dim oCtl As Control
    Dim oBase As Control
    Set oBase = Me.Controls(sBase)
    Dim i As Integer
    On Error Resume Next
    For i = LBound(aControls) To UBound(aControls)
        Set oCtl = Controls(aControls(i))
        oCtl.Left = oBase.Left
        oCtl.Width = oBase.Width
        oCtl.Top = oBase.Top
        oCtl.Height = oBase.Height
    Next
End Sub

Private Function Validate(ctl As Control) As Boolean
    Dim xProp As ScriptProperty
    Set xProp = PropSet(lvProps.SelectedItem.Tag)
    
    If Len(xProp.Pattern) = 0 Then
        Validate = True
        Exit Function
    End If
    
    Dim sTmp As String
    Dim sExpr As String
    With lvProps.SelectedItem
        Select Case xProp.DataType
        Case ptString
            sTmp = txtTextValue
        Case ptInteger
            sTmp = txtNumValue
        Case ptBoolean
            sTmp = IIf(chkBoolValue.Value = vbChecked, sTrue, sFalse)
        Case ptDate
            sTmp = Format(dtDateValue, "Short Date", vbUseSystemDayOfWeek, vbUseSystem)
        Case ptFloat
            sTmp = txtFloatValue
        Case ptChoice
            sTmp = cbChoiceValue
        Case ptCurrency
            sTmp = Left$(cbChoiceValue, 3)
        End Select
'        .SubItems(1) = PropValString(xProp)
    End With

    If Left$(xProp.Pattern, 1) = "=" Then
        sExpr = Mid$(xProp.Pattern, 2)
'        sExpr = Replace(Mid$(xProp.Pattern, 2), "%1", CStr(sTmp))
        Validate = ScriptDoValidate(ActiveConfigModule(), sExpr, sTmp)
    Else
        Dim re As New RegExp
        With re
            .Global = False
            .IgnoreCase = True
            .MultiLine = False
            .Pattern = xProp.Pattern
            Validate = .Test(sTmp)
        End With
    End If
    Dim sMsg As String
    If Not Validate Then
        sMsg = xProp.ValidationFailMessage
        If Len(sMsg) = 0 Then
            sMsg = ScriptValidationMessage(ActiveConfigModule())
        End If
        If Len(sMsg) > 0 Then
            sMsg = Replace$(sMsg, "%1", Replace(sTmp, """", """"""))
        Else
            sMsg = LoadResStringLEx(1811, sTmp, xProp.Pattern)
            If Len(sMsg) = 0 Then
                sMsg = "Value '" & sTmp & "' does not match pattern '" & xProp.Pattern & "'"
            End If
        End If
        MyMsgBox sMsg, vbInformation + vbOKOnly, LoadResStringL(1812)
    End If
End Function

Private Sub txtTextValue_Validate(Cancel As Boolean)
    Cancel = Not Validate(txtTextValue)
End Sub
