VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmScriptWiz 
   Caption         =   "MT2OFX Script Builder"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   480
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmFields 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5055
      Left            =   1080
      TabIndex        =   6
      Top             =   1800
      Width           =   10095
      Begin VB.TextBox txtPattern 
         Height          =   375
         Left            =   1320
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   3840
         Width           =   2775
      End
      Begin VB.TextBox txtHeader 
         Height          =   375
         Left            =   1320
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   3120
         Width           =   2775
      End
      Begin VB.ComboBox cbField 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3480
         Width           =   2775
      End
      Begin VB.ComboBox cbHeaderLine 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   120
         Width           =   9975
      End
      Begin MSComctlLib.ListView lvFields 
         Height          =   2535
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Col"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Header"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Field"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Pattern"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Selected Field:"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   1215
      End
   End
   Begin VB.Frame frmParameters 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4695
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   9015
      Begin VB.ComboBox cbDelimiter 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7560
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   0
         Width           =   1455
      End
      Begin VB.ComboBox cbCodePage 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   0
         Width           =   5295
      End
      Begin MSComctlLib.ListView lvParams 
         Height          =   2775
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Value"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Field delimiter:"
         Height          =   255
         Left            =   6360
         TabIndex        =   12
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Code Page:"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   975
      End
   End
   Begin MSComctlLib.TabStrip tsTabs 
      Height          =   5415
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9551
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Parameters"
            Key             =   "params"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fields"
            Key             =   "fields"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtFile 
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Text            =   "CSV File goes here"
      Top             =   120
      Width           =   10815
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   11040
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   11040
      TabIndex        =   0
      Top             =   6480
      Width           =   855
   End
End
Attribute VB_Name = "frmScriptWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const csTabParams As String = "params"
Const csTabFields As String = "fields"
Const ciNumHeaderLines As Long = 30

Dim xInputFile As InputFile

Dim iCurFrame As Integer
Dim frmCurFrame As Frame

Private Declare Function GetACP Lib "kernel32" () As Long

Private Sub cbHeaderLine_Click()
    Dim aFields, f
    Dim i As Long
    Dim sTmp As String
    Dim sDelim As String
    sDelim = Left$(cbDelimiter, 1)
    If sDelim = " " Then sDelim = vbTab
    sTmp = cbHeaderLine
    aFields = txtParseLineDelimited(sTmp, sDelim, False)
    With lvFields
        .ListItems.Clear
        For i = LBound(aFields) To UBound(aFields)
            With .ListItems.Add(, , CStr(i))
                .Tag = fldSkip
                .SubItems(1) = aFields(i)
                .SubItems(2) = "Skip"
            End With
        Next
    End With
    
End Sub

Private Sub cmdBrowse_Click()
    With cdOpen
        .FileName = txtFile
        .Filter = "All Files (*.*)|*.*"
        .FilterIndex = 1
        .DialogTitle = "Choose Downloaded File"
        .ShowOpen
        If Len(.FileName) <> 0 Then
            txtFile = .FileName
            Call OpenFile
        End If
    End With
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub Form_Load()
    Dim aiCodePages() As Long
    Dim i As Long, iTmp As Long
    Dim sTmp As String, sTmp2 As String
    
    txtFile = ""
    With tsTabs
        frmParameters.Left = .ClientLeft
        frmParameters.Width = .ClientWidth
        frmParameters.Top = .ClientTop
        frmParameters.Height = .ClientHeight
        lvParams.Width = frmParameters.Width - (frmParameters.Left * 2)
        frmFields.Left = .ClientLeft
        frmFields.Width = .ClientWidth
        lvFields.Width = frmFields.Width - (frmFields.Left * 2)
        cbHeaderLine.Width = frmFields.Width - (frmFields.Left * 2)
        frmFields.Top = .ClientTop
        frmFields.Height = .ClientHeight
    End With
    Set tsTabs.SelectedItem = tsTabs.Tabs(csTabParams)
    With cbDelimiter
        .Clear
        .AddItem ", (Comma)"
        .AddItem "; (Semicolon)"
        .AddItem "  (Tab)"
        sTmp = GetLocaleString(LOCALE_USER_DEFAULT, LOCALE_SLIST, ",")
        .ListIndex = FindInCombo(cbDelimiter, sTmp)
    End With
    
    With cbCodePage
        .Clear
        .AddItem "      " & "(System setting)"
        aiCodePages = GetCodePageList()
        For i = LBound(aiCodePages) To UBound(aiCodePages)
            iTmp = aiCodePages(i)
            If IsValidCodePage(iTmp) And iTmp <> CP_USER Then
                sTmp = GetCodePageName(iTmp)
                If Left$(sTmp, 1) <> "_" Then
                    sTmp2 = Right$("     " & CStr(iTmp), 5)
                    .AddItem sTmp2 & " : " & sTmp
                End If
            End If
        Next
        .ListIndex = 0
    End With
    
    With cbField
        .Clear
        .AddItem "Skip": .ItemData(.NewIndex) = fldSkip
        .AddItem "Account Num": .ItemData(.NewIndex) = fldAccountNum
        .AddItem "Currency": .ItemData(.NewIndex) = fldCurrency
        .AddItem "Closing Balance": .ItemData(.NewIndex) = fldClosingBal
        .AddItem "Avail Balance": .ItemData(.NewIndex) = fldAvailBal
        .AddItem "Book Date": .ItemData(.NewIndex) = fldBookDate
        .AddItem "Value Date": .ItemData(.NewIndex) = fldValueDate
        .AddItem "Amount Credit": .ItemData(.NewIndex) = fldAmtCredit
        .AddItem "Amount Debit": .ItemData(.NewIndex) = fldAmtDebit
        .AddItem "Memo": .ItemData(.NewIndex) = fldMemo
        .AddItem "Balance Date": .ItemData(.NewIndex) = fldBalanceDate
        .AddItem "Amount": .ItemData(.NewIndex) = fldAmount
        .AddItem "Payee": .ItemData(.NewIndex) = fldPayee
        .AddItem "Txn Date": .ItemData(.NewIndex) = fldTransactionDate
        .AddItem "Txn Time": .ItemData(.NewIndex) = fldTransactionTime
        .AddItem "Cheque Num": .ItemData(.NewIndex) = fldChequeNum
        .AddItem "FITID": .ItemData(.NewIndex) = fldFITID
        .AddItem "Empty": .ItemData(.NewIndex) = fldEmpty
        .AddItem "Branch ID": .ItemData(.NewIndex) = fldBranch
        .AddItem "Amount Sign": .ItemData(.NewIndex) = fldSign
        .AddItem "Category": .ItemData(.NewIndex) = fldCategory
        .AddItem "P. City": .ItemData(.NewIndex) = fldPayeeCity
        .AddItem "P. State": .ItemData(.NewIndex) = fldPayeeState
        .AddItem "P. Postcode": .ItemData(.NewIndex) = fldPayeeZip
        .AddItem "P. Country": .ItemData(.NewIndex) = fldPayeeCountry
        .AddItem "P. Phone": .ItemData(.NewIndex) = fldPayeePhone
        .AddItem "P. Address 1": .ItemData(.NewIndex) = fldPayeeAddress1
        .AddItem "P. Address 2": .ItemData(.NewIndex) = fldPayeeAddress2
        .AddItem "P. Address 3": .ItemData(.NewIndex) = fldPayeeAddress3
        .AddItem "P. Address 4": .ItemData(.NewIndex) = fldPayeeAddress4
        .AddItem "P. Address 5": .ItemData(.NewIndex) = fldPayeeAddress5
        .AddItem "P. Address 6": .ItemData(.NewIndex) = fldPayeeAddress6
    End With
End Sub

Private Sub lvFields_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim i As Long
    txtHeader = Item.SubItems(1)
    For i = 0 To cbField.ListCount - 1
        If cbField.ItemData(i) = Item.Tag Then
            cbField.ListIndex = i
            Exit For
        End If
    Next
    txtPattern = Item.SubItems(3)
End Sub

Private Sub tsTabs_Click()
    Dim newframe As Frame
    If tsTabs.SelectedItem.Index = iCurFrame _
        Then Exit Sub ' No need to change frame.
    ' Otherwise, hide old frame, show new.
    Select Case tsTabs.SelectedItem.Key
    Case csTabParams
        Set newframe = frmParameters
    Case csTabFields
        Set newframe = frmFields
    End Select
    newframe.Visible = True
    If Not frmCurFrame Is Nothing Then frmCurFrame.Visible = False
    ' Set mintCurFrame to new value.
    iCurFrame = tsTabs.SelectedItem.Index
    Set frmCurFrame = newframe
End Sub

Private Sub OpenFile()
    Dim i As Long
    Dim sLine As String
    Dim sTmp As String
    
    cbHeaderLine.Clear
    lvFields.ListItems.Clear
    lvParams.ListItems.Clear
    Set xInputFile = New InputFile
    sTmp = Trim(Left(cbCodePage, 5))
    If Len(sTmp) = 0 Then
        i = GetACP()
    Else
        i = CLng(sTmp)
    End If
    xInputFile.CodePage = i
    If xInputFile.OpenFile(txtFile) Then
        For i = 1 To ciNumHeaderLines
            If xInputFile.AtEOF Then
                Exit For
            End If
            sLine = xInputFile.ReadLine
            cbHeaderLine.AddItem sLine
        Next
        cbHeaderLine.ListIndex = 0
    Else
        MsgBox "Unable to open file."
    End If
End Sub
