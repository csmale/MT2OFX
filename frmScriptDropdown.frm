VERSION 5.00
Begin VB.Form frmScriptDropdown 
   Caption         =   "Scripted Selection"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   4005
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRemove 
      Cancel          =   -1  'True
      Caption         =   "Remove"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.ComboBox cbCombo 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Text            =   "cbCombo"
      Top             =   720
      Width           =   3615
   End
   Begin VB.ComboBox cbChoice 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   3615
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblDescription 
      Caption         =   "Label1"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmScriptDropdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frmScriptDropdown
'    Project    : MT2OFX
'
'    Description: Form to support script function ChooseFromList
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit

Public Choices As Variant
Public FixedList As Boolean
Public Selection As String
Public ValueList As String
Public HelpText As String
Public DialogTitle As String
Public NewChoices As String
Public ValidateProc As String
Public ListDelimiter As String

Private Const csDelim As String = "\"

' The form has two comboboxes - one with type vbComboDropdownList (user cannot add own value) and one
' with type vbComboDropdown (user can add own value if required). As the style cannot be changed at
' runtime, there are two controls- one is hidden, the other is stored in myCombo.
Private myCombo As ComboBox
Private bChanged As Boolean


'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cbChoice_Click
' Description:       Click event handler
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       18/03/2011-13:50:14
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cbChoice_Click()
    Call HandleClickEvent
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cbCombo_Change
' Description:
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       9/28/2005-22:05:22
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cbCombo_Change()
    Call HandleClickEvent
    bChanged = True
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cbCombo_Click
' Description:       Click event handler
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       18/03/2011-13:50:36
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cbCombo_Click()
    Call HandleClickEvent
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdCancel_Click
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       9/28/2005-22:06:23
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdCancel_Click()
    Selection = ""
    Unload Me
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdOK_Click
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       9/28/2005-22:06:38
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdOK_Click()
    Dim sNewList, i, sItem
    Selection = myCombo.Text
    If Len(ValidateProc) > 0 Then
        If Not ScriptDoValidate(ActiveScriptModule(), ValidateProc, Selection) Then
            MsgBox ScriptValidationMessage(ActiveScriptModule()), vbExclamation + vbOKOnly, "Invalid entry"
            Exit Sub
        End If
    End If
    If Not FixedList Then
        sNewList = ""
        For i = 0 To myCombo.ListCount - 1
            sItem = myCombo.List(i)
            If Len(sNewList) = 0 Then
                sNewList = sItem
            Else
                sNewList = sNewList & ListDelimiter & sItem
            End If
        Next
        If bChanged Then
            sItem = Selection
            If Len(sNewList) = 0 Then
                sNewList = sItem
            Else
                sNewList = sNewList & ListDelimiter & sItem
            End If
        End If
        NewChoices = sNewList
    End If
    Unload Me
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cmdRemove_Click
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       9/28/2005-22:49:47
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdRemove_Click()
    Dim i As Integer
    Dim sItem As String
    i = myCombo.ListIndex
    sItem = myCombo.List(i)
    myCombo.RemoveItem i
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Form_Load
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       9/28/2005-22:06:51
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()
    cbCombo.Left = cbChoice.Left
    cbCombo.Top = cbChoice.Top
    cbCombo.Width = cbChoice.Width
'    cbCombo.Height = cbChoice.Height
    If FixedList Then
        Set myCombo = cbChoice
        cbCombo.Visible = False
        cbCombo.Enabled = False
        cmdRemove.Visible = False
        cmdRemove.Enabled = False
    Else
        Set myCombo = cbCombo
        cbChoice.Visible = False
        cbChoice.Enabled = False
    End If
    myCombo.Enabled = True
    myCombo.Visible = True
    myCombo.Clear
    If Len(ListDelimiter) = 0 Then
        ListDelimiter = csDelim
    ElseIf Len(ListDelimiter) > 1 Then
        ListDelimiter = Left$(ListDelimiter, 1)
    End If
    Dim aList As Variant
    aList = Split(ValueList, ListDelimiter)
    Dim i
    If IsArray(aList) Then
        For i = LBound(aList) To UBound(aList)
            myCombo.AddItem aList(i)
        Next
    End If
    If Len(Selection) > 0 Then
        myCombo.ListIndex = FindInCombo(myCombo, Selection)
    End If
    Me.lblDescription.Caption = HelpText
    If Len(DialogTitle) > 0 Then
        Me.Caption = DialogTitle
    End If
    bChanged = False
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       AddToChoiceList
' Description:       Add a string to NewChoices if not already present
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       9/28/2005-22:08:46
'
' Parameters :       sNewChoice (String)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub AddToChoiceList(sNewChoice As String)
    Dim aList() As String
    Dim i As Long
    aList = Split(NewChoices, csDelim)
    For i = LBound(aList) To UBound(aList)
        If sNewChoice = aList(i) Then
            Exit Sub
        End If
    Next
    If Len(NewChoices) > 0 Then
        NewChoices = NewChoices & ListDelimiter
    End If
    NewChoices = NewChoices & sNewChoice
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       RemoveFromChoiceList
' Description:       Remove a string from NewChoices if it's there
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       9/28/2005-22:09:18
'
' Parameters :       sChoice (String)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub RemoveFromChoiceList(sChoice As String)
    Dim aList() As String
    Dim i As Long
    Dim iFound As Long
    iFound = -1
    aList = Split(NewChoices, ListDelimiter)
    For i = LBound(aList) To UBound(aList)
        If sChoice = aList(i) Then
            iFound = i
            Exit For
        End If
    Next
    If iFound < 0 Then
        Exit Sub
    End If
    For i = iFound To UBound(aList) - 1
        aList(i) = aList(i + 1)
    Next
    ReDim Preserve aList(0 To UBound(aList) - 1)
    NewChoices = Join(aList, csDelim)
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       HandleClickEvent
' Description:       [type_description_here]
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       18/03/2011-13:51:04
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub HandleClickEvent()
    Dim i As Integer
    i = myCombo.ListIndex
    If i >= 0 Then
        cmdRemove.Enabled = Not FixedList
    Else
        cmdRemove.Enabled = FixedList
    End If
    cmdOK.Enabled = (Len(myCombo.Text) > 0)
End Sub
