VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQuickStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Script QuickStart"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkExcludeOthers 
      Caption         =   "Leave scripts from other countries out of configuration"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   5160
      WhatsThisHelpID =   1904
      Width           =   5175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   5160
      WhatsThisHelpID =   1905
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Top             =   5160
      WhatsThisHelpID =   1906
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvBanks 
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      WhatsThisHelpID =   1903
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Bank"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Format"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Script"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.ComboBox cbCountry 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Description"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      WhatsThisHelpID =   1901
      Width           =   7695
   End
   Begin VB.Label lblCountry 
      Caption         =   "Country"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      WhatsThisHelpID =   1902
      Width           =   1455
   End
End
Attribute VB_Name = "frmQuickStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : frmQuickStart
'    Project    : MT2OFX
'
'    Description: Quick Start form to help with Script Selection
'
'    Modified   : $Author: Colin $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/frmQuickStart.frm 6     24/11/09 22:04 Colin $"
' $History: frmQuickStart.frm $
' 
' *****************  Version 6  *****************
' User: Colin        Date: 24/11/09   Time: 22:04
' Updated in $/MT2OFX
' for 3.6 beta
'
' *****************  Version 5  *****************
' User: Colin        Date: 30/08/09   Time: 13:20
' Updated in $/MT2OFX
'
' *****************  Version 4  *****************
' User: Colin        Date: 25/11/08   Time: 22:22
' Updated in $/MT2OFX
' moving vss server!
'
' *****************  Version 2  *****************
' User: Colin        Date: 15/04/08   Time: 22:28
' Updated in $/MT2OFX
' First code complete

'</CSCC>

Dim xDoc As MSXML2.DOMDocument30
Dim xCountry As MSXML2.IXMLDOMNode
Dim xBankList As MSXML2.IXMLDOMNodeList

Const csScriptList As String = "scriptcat.xml"
Const csxAllCountries As String = "/mt2ofx/region"
Const csxCountryById As String = "/mt2ofx/region[@id='%1']"
Const csxAllScripts As String = "/mt2ofx/bankscript"
Const csxScriptsByRegion As String = "/mt2ofx/bankscript[region/text()='%1' or region/text()='zz']"
Const csxScriptById As String = "/mt2ofx/bankscript[@id='%1']"
Const csxFormatById As String = "/mt2ofx/format[@id='%1']"
Const csxLanguageById As String = "/mt2ofx/language[text()='%1']"
Const csxLanguageFirst As String = "/mt2ofx/language[0]"
Const csRegionAllScripts As String = "--"
Const csRegionGeneric As String = "zz"
Const csLanguageEnglish As String = "en"
Const csFormatMT940 As String = "MT940"

Dim sXml As String
Dim asCountries(100) As String
Dim iCountries As Integer
Dim sLang As String
Dim sLocation As String

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       cbCountry_Click
' Description:       Selection of a country (or territory or region or whatever)
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       1/10/2007-00:07:23
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cbCountry_Click()
    Dim sCC As String
    If cbCountry.ListIndex < 0 Then Exit Sub
    sCC = asCountries(cbCountry.ItemData(cbCountry.ListIndex))
    Me.lvBanks.ListItems.Clear
' capture currently selected region/country
    Set xCountry = xDoc.selectSingleNode(Replace(csxCountryById, "%1", sCC))
' get list of banks with this country in their definition
    If sCC = csRegionAllScripts Then
        Set xBankList = xDoc.selectNodes(csxAllScripts)
    Else
        Set xBankList = xDoc.selectNodes(Replace(csxScriptsByRegion, "%1", sCC))
    End If
    
    Dim xBankScript As IXMLDOMNode
    Dim xBank As IXMLDOMNode
    Dim xItem As ListItem
    Dim sBankCode As String
    Dim sScript As String
    Dim sFormat As String
    Dim xFormat As IXMLDOMNode
    Dim xFormat2 As IXMLDOMNode
    Dim xScript As IXMLDOMNode
' iterate through the selected bank scripts
    For Each xBankScript In xBankList
        Set xItem = lvBanks.ListItems.Add(, xBankScript.selectSingleNode("@id").Text, xBankScript.selectSingleNode("@name").Text)
        sFormat = xBankScript.selectSingleNode("@format").Text
        Set xFormat = xDoc.selectSingleNode(Replace$(csxFormatById, "%1", sFormat))
        If Not xFormat Is Nothing Then
            Set xFormat2 = xFormat.selectSingleNode("@" & sLang)
            If Not xFormat2 Is Nothing Then
                sFormat = xFormat2.Text
            End If
        End If
        Set xScript = xBankScript.selectSingleNode("script")
        If xScript Is Nothing Then
            sScript = ""
        Else
            sScript = xScript.Text
        End If
        xItem.SubItems(1) = sFormat
        xItem.SubItems(2) = sScript
        xItem.Checked = True
    Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim iAns As Integer
    iAns = MyMsgBox(LoadResStringL(1909), vbYesNoCancel, "MT2OFX")
    Select Case iAns
    Case vbYes
        SaveNewScripts
        Unload Me
    Case vbNo
        Unload Me
    End Select
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Form_Load
' Description:       Load Event
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       1/10/2007-00:08:04
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()
    Dim iThisCountry As Integer
    Dim xCountries As IXMLDOMNodeList
    Dim xLanguage As IXMLDOMNode
    Dim sTmp As String

    LocaliseForm Me, 1900
    
    sXml = Cfg.AppPath & "\" & csScriptList
' the language resource files contain a string (#1907) which indicates a language code to be used for
' the presentation of the script information. Like this additional language resource files can point
' to the "least worst" of the languages in the script.xml file. This is checked against the languages present
' in the XML file
    sLang = LoadResStringL(1907)
    
' get the country from the user's location (control panel Regional Options)
' note: this is NOT the locale country (language/culture)
    sLocation = GetUserCountry()
    LogMessage False, True, "Resource language: " & sLang & ", User country: " & sLocation

    Me.cbCountry.Clear
    
    Set xDoc = New DOMDocument30
    xDoc.async = False
    If Not xDoc.Load(sXml) Then
        sTmp = LoadResStringLEx(1908, xDoc.parseError.reason)
        LogMessage True, True, sTmp, , vbOKOnly
        Exit Sub
    End If

' pick a language for the script data
    Set xLanguage = xDoc.selectSingleNode(Replace$(csxLanguageById, "%1", sLang))
    If xLanguage Is Nothing Then
        Set xLanguage = xDoc.selectSingleNode(csxLanguageFirst)
    End If
    If xLanguage Is Nothing Then
        sLang = csLanguageEnglish
    Else
        sLang = xLanguage.Text
    End If
    
    Set xCountries = xDoc.selectNodes(csxAllCountries)
    iThisCountry = -1
    iCountries = 0
    On Error Resume Next
    Me.cbCountry.ListIndex = -1
    For Each xCountry In xCountries
        Me.cbCountry.AddItem xCountry.selectSingleNode("@" & sLang).Text
        iCountries = iCountries + 1
        asCountries(iCountries) = xCountry.selectSingleNode("@id").Text
        Me.cbCountry.ItemData(Me.cbCountry.newIndex) = iCountries
        If asCountries(iCountries) = sLocation Then
            iThisCountry = Me.cbCountry.newIndex
            cbCountry.ListIndex = iThisCountry
        End If
    Next
    Set xCountry = Nothing
'    cbCountry.ListIndex = iThisCountry
End Sub

Private Sub SaveNewScripts()
    Dim i As Integer
    Dim iScript
    Dim xScript As IXMLDOMNode
    Dim xItem As ListItem
    Dim xDic As New Scripting.Dictionary
    Dim sScript As String
    Dim sFormat As String
    Dim sBank As String
    Dim iBankRule As Integer
    Dim xSection As IXMLDOMNode
    Dim xMatches As IXMLDOMNodeList
    Dim xMatch As IXMLDOMNode
    Dim xBankCfg As BankConfig
    Dim iLines As Integer
    Dim sPattern As String
    
' clear current script config
    ClearMySection IniSectionText
    ClearMySection IniSectionBankRules
    
' put selected scripts in first, activated according to the checkbox
    iScript = 0
    iBankRule = 0
    For Each xItem In lvBanks.ListItems
        sFormat = xItem.SubItems(1)
        If sFormat = csFormatMT940 Then
            Set xScript = xDoc.selectSingleNode(Replace(csxScriptById, "%1", xItem.Key))
            If xScript Is Nothing Then GoTo nextitem
            Set xSection = xScript.selectSingleNode("section")
            If xSection Is Nothing Then GoTo nextitem
            sBank = xSection.selectSingleNode("@name").Text
            ClearMySection sBank
            Set xBankCfg = New BankConfig
            xBankCfg.BankKey = sBank
            xBankCfg.IDString = xItem.Text
            xBankCfg.Structured86 = (xSection.selectSingleNode("structured86").Text = "1")
            xBankCfg.SkipEmptyMemoFields = (xSection.selectSingleNode("skipemptymemofields").Text = "1")
            xBankCfg.ScriptFile = xItem.SubItems(2)
            xBankCfg.Save sBank
            Set xMatches = xSection.selectNodes("match")
            For Each xMatch In xMatches
                iBankRule = iBankRule + 1
                iLines = CInt(xMatch.selectSingleNode("lines").Text)
                sPattern = xMatch.selectSingleNode("pattern").Text
                PutMyString IniSectionBankRules, "Bank" & CStr(iBankRule), sBank & "," & CStr(iLines) & "," & sPattern
            Next
        Else
            iScript = iScript + 1
            sScript = xItem.SubItems(2)
            ' add to section - activated if checked in listbox
            PutMyString IniSectionText, "Script" & CStr(iScript), IIf(xItem.Checked, "1", "0") & "," & sScript
        End If
        ' remember this script to avoid duplicates
        xDic.Item(sScript) = "1"
nextitem:
    Next
' now put all other scripts in the list, not activated
' only same country or all countries?
    If chkExcludeOthers.Value = vbUnchecked Then
        For Each xScript In xDoc.selectNodes(csxAllScripts)
            sScript = xScript.selectSingleNode("script").Text
            sFormat = xScript.selectSingleNode("@format").Text
            If Not xDic.Exists(sScript) Then
                If sFormat = csFormatMT940 Then
                Else
                    iScript = iScript + 1
                    PutMyString IniSectionText, "Script" & CStr(iScript), "0," & sScript
                End If
            End If
        Next
    End If
End Sub
