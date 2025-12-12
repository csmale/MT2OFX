VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScriptEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Session"
Attribute VB_Ext_KEY = "Member1" ,"Statement"
Attribute VB_Ext_KEY = "Member2" ,"Statement"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member3" ,"Statement"
Attribute VB_Ext_KEY = "Member4" ,"Txn"
Attribute VB_Ext_KEY = "Member5" ,"Txn"
Attribute VB_Ext_KEY = "Member6" ,"Txn"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : ScriptEnv
'    Project    : MT2OFX
'
'    Description:
'
'    Modified   : $Author: Colin $ $Date: 27/03/11 23:33 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/ScriptEnv.cls 29    27/03/11 23:33 Colin $"
' $History: ScriptEnv.cls $
' 
' *****************  Version 29  *****************
' User: Colin        Date: 27/03/11   Time: 23:33
' Updated in $/MT2OFX
' multiple account support
'
' *****************  Version 28  *****************
' User: Colin        Date: 14/11/10   Time: 23:52
' Updated in $/MT2OFX
'
' *****************  Version 27  *****************
' User: Colin        Date: 9/01/10    Time: 23:22
' Updated in $/MT2OFX
' add SetStatement()
'
' *****************  Version 26  *****************
' User: Colin        Date: 15/06/09   Time: 19:25
' Updated in $/MT2OFX
' For transfer to new laptop
'
' *****************  Version 26  *****************
' User: Colin        Date: 17/01/09   Time: 23:04
' Updated in $/MT2OFX
'
' *****************  Version 25  *****************
' User: Colin        Date: 25/11/08   Time: 22:24
' Updated in $/MT2OFX
' moving vss server!
'
' *****************  Version 23  *****************
' User: Colin        Date: 19/04/08   Time: 22:37
' Updated in $/MT2OFX
' add HTML clipboard support
' add unified input support - ProcessAsMt940()
'
' *****************  Version 22  *****************
' User: Colin        Date: 7/12/06    Time: 15:07
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 21  *****************
' User: Colin        Date: 25/04/06   Time: 21:45
' Updated in $/MT2OFX
'
' *****************  Version 19  *****************
' User: Colin        Date: 1/03/06    Time: 23:00
' Updated in $/MT2OFX
'
' *****************  Version 17  *****************
' User: Colin        Date: 2/11/05    Time: 23:03
' Updated in $/MT2OFX
' V3.4 beta 1
'
' *****************  Version 16  *****************
' User: Colin        Date: 11/06/05   Time: 19:33
' Updated in $/MT2OFX
'
' *****************  Version 15  *****************
' User: Colin        Date: 6/03/05    Time: 23:42
' Updated in $/MT2OFX
'</CSCC>

' This class defines all the objects etc. exposed to scripting environment.
' Access is through property functions, with the set/let functions defined
' as Friend so as to make them inaccessible via scripting. Effectively
' the items become "read only" in the scripting environment.

#Const debugmode = 0

Public Enum Str86Fields
    sfBookingText = 0
    sfBatchNum = 10
    sfDetails0 = 20
    sfDetailsLast = 29
    sfBankPayee = 30
    sfAcctPayee = 31
    sfNamePayee = 32
    sfNamePayee2 = 33
    sfTextCodeSupplement = 34
    sfExtra0 = 60
    sfExtraLast = 63
End Enum

Public AbortRequested As Boolean
Private xStatement As Statement
Private xTxn As Txn
Private xCfg As ProgramConfig
Private xBcfg As BankConfig
Private xSession As Session
Private xConfigForm As frmScriptCfg
Private xScriptProperties As ScriptPropertySet
Private xHTML As MSHTML.IHTMLDocument3
Private xLastLine As String
Private sConfigMod As String

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ConfigForm
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12/11/2004-16:42:56
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get ConfigForm() As frmScriptCfg
    If xConfigForm Is Nothing Then
        Set xConfigForm = New frmScriptCfg
    End If
    Set ConfigForm = xConfigForm
End Property

Public Property Get Session() As Session
    Set Session = xSession
End Property

Friend Property Set Session(X As Session)
    Set xSession = X
End Property

Public Property Get Statement() As Statement
    Set Statement = xStatement
End Property

Friend Property Set Statement(X As Statement)
    Set xStatement = X
End Property

Public Sub SetStatement(X As Statement)
    Set xStatement = X
End Sub

Public Property Get Txn() As Txn
    Set Txn = xTxn
End Property

Friend Property Set Txn(X As Txn)
    Set xTxn = X
End Property

Public Property Get Cfg() As ProgramConfig
    Set Cfg = xCfg
End Property

Friend Property Set Cfg(X As ProgramConfig)
    Set xCfg = X
End Property

Public Property Get Bcfg() As BankConfig
    Set Bcfg = xBcfg
End Property

Friend Property Set Bcfg(X As BankConfig)
    Set xBcfg = X
End Property

Public Function MakeGUID() As String
    MakeGUID = GenerateGUID()
End Function

Public Sub Message(ToScreen As Variant, ToFile As Variant, Text As Variant, Title As Variant)
    LogMessage CBool(ToScreen), CBool(ToFile), CStr(Text), CStr(Title)
End Sub

Public Function GetConfigString(sSection As Variant, sKey As Variant, sDefault As Variant) As String
    GetConfigString = GetMyString(CStr(sSection), CStr(sKey), CStr(sDefault))
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetConfigStringEx
' Description:       Read an Ini string from a given file
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       29/03/2004-08:22:35
'
' Parameters :       sSection (String)
'                    sKey (String)
'                    sDefault (String)
'                    SFile (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetConfigStringEx(sSection As Variant, sKey As Variant, sDefault As Variant, sFile As Variant) As String
    GetConfigStringEx = ReadIniString(CStr(sFile), CStr(sSection), CStr(sKey), CStr(sDefault))
End Function

Public Function Version() As String
    Version = CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
End Function

' additions for text file support
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ReadLine
' Description:       Returns the next line from the input file
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       23/11/2003-22:04:54
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ReadLine() As String
    Dim sTmp As String
    On Error GoTo baleout
    sTmp = Session.InputFile.ReadLine
    xLastLine = sTmp
    ShowProgress Session.InputFile.Pos
    ReadLine = sTmp
    Exit Function
baleout:
    LogMessage False, True, "Error in ReadLine(): " & Err.Description, , vbCritical + vbOKOnly
    Err.Raise Err.Number, "ReadLine", Err.Description, Err.HelpFile, Err.HelpContext
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       LastLine
' Description:       Return the last line returned by ReadLine (again!)
'                    Useful for debugging
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       4/11/2006-21:38:24
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function LastLine() As String
    LastLine = xLastLine
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       EntireFile
' Description:       Returns the entire file as a long string
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       7/25/2005-18:23:53
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function EntireFile() As String
    Dim sTmp As String
    On Error GoTo baleout
    With Session.InputFile
        .Rewind
        sTmp = .EntireFile
        ShowProgress .Length
    End With
    EntireFile = sTmp
    Exit Function
baleout:
    LogMessage False, True, "Error in EntireFile(): " & Err.Description, , vbCritical + vbOKOnly
    Err.Raise Err.Number, "EntireFile", Err.Description, Err.HelpFile, Err.HelpContext
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       EntireFileHTML
' Description:       Returns the entire file but the HTML clipboard intead of the text clipboard
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       4/19/2008-22:35:28
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function EntireFileHTML() As String
    Dim sTmp As String
    sTmp = HTMLClipboardData
    If Len(sTmp) = 0 Then
        sTmp = EntireFile()
    Else
        ShowProgress Session.InputFile.Length
    End If
    EntireFileHTML = sTmp
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       LoadAsMt940
' Description:       Entry point from LoadTextFile for unified input support
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       4/19/2008-22:34:25
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function LoadAsMt940() As Boolean
    LoadAsMt940 = ProcessAsMt940()
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       AtEOF
' Description:       Returns True if at end of input file
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       23/11/2003-22:05:39
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function AtEOF() As Boolean
    AtEOF = Session.InputFile.AtEOF
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Rewind
' Description:       Rewinds input file
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       15 Dec 2004-12:46:25
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub Rewind()
    Session.InputFile.Rewind
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ParseLineDelimited
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       23/11/2003-22:11:26
'
' Parameters :       Line (String)
'                    Delimiter (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ParseLineDelimited(ByVal Line As Variant, ByVal Delimiter As Variant, Optional ByVal UseBackslash As Boolean = True) As Variant
    ParseLineDelimited = txtParseLineDelimited(CStr(Line), CStr(Delimiter), UseBackslash)
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ParseLineFixed
' Description:       Split a line at fixed boundaries
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       23/11/2003-22:12:39
'
' Parameters :       Line (String)
'                    Pattern (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ParseLineFixed(ByVal Line As Variant, ByVal Pattern As Variant) As Variant
#If debugmode > 0 Then
    LogMessage False, True, "Enter: ParseLineFixed"
#End If
    ParseLineFixed = txtParseLineFixed(Line, Pattern)
#If debugmode > 0 Then
    LogMessage False, True, "Leave: ParseLineFixed"
#End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       NewStatement
' Description:       Creates a new Statement for a text script
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       06/12/2003-23:13:08
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function NewStatement() As Statement
    Dim oTmp As New Statement
    Debug.Assert Not (Session Is Nothing)
    If Session Is Nothing Then
        LogMessage False, True, "NewStatement called with no current Session"
        Err.Raise vbObjectError + 1, , "NewStatement called with no current Session"
        Abort
        Set NewStatement = Nothing
        Exit Function
    End If
    
    Session.Statements.Add oTmp
    Set oTmp.Session = Session
    Set xStatement = oTmp
    
    Set NewStatement = oTmp
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       NewTransaction
' Description:       Creates a new transaction object for text script
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       06/12/2003-23:15:33
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function NewTransaction() As Txn
    Dim oTmp As New Txn
    Debug.Assert Not (xStatement Is Nothing)
    If xStatement Is Nothing Then
        LogMessage False, True, "NewTransaction called with no current Statement"
        Err.Raise vbObjectError + 1, , "NewTransaction called with no current Statement"
        Abort
        Set NewTransaction = Nothing
        Exit Function
    End If
    Set oTmp.Statement = xStatement
    oTmp.Index = xStatement.Txns.Count + 1
    xStatement.Txns.Add oTmp
    Set xTxn = oTmp
    Set NewTransaction = oTmp
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Abort
' Description:       Remembers error flag to terminate processing
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       07/12/2003-17:12:01
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub Abort()
    AbortRequested = True
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DecimalSeparator
' Description:       Returns the system's configured decimal separator
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       20/01/2004-15:25:22
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function DecimalSeparator() As String
    DecimalSeparator = SystemDecimalSeparator()
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       UserCurrency
' Description:       Returns user's default currency (ISO code)
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       03/02/2005-10:55:26
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function UserCurrency() As String
    UserCurrency = GetUserCurrency()
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       UserDateSequence
' Description:       Return string representing user's default date sequence (MDY, DMY, YMD)
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/3/2008-17:39:05
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function UserDateSequence() As String
    UserDateSequence = GetUserDateSequence()
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       UserLocale
' Description:       Returns LCID
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       22/04/2010-23:08:27
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function UserLocale() As Long
    UserLocale = GetProgLocale()
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       UserLanguage
' Description:       Returns user language ISO639 code e.g. "en" for English, "nl" for Dutch
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       22/04/2010-23:41:16
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function UserLanguage() As String
    UserLanguage = GetUserLanguage()
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ParseNumber
' Description:       Parse a string into a number, using the given decimal sign
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       29/02/2004-22:28:10
'
' Parameters :       s (String)
'                    dec (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ParseNumber(vs As Variant, dec As Variant) As Double
    Dim s As String
    s = CStr(vs)
    ParseNumber = NumberFromString(s, CStr(dec))
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Application
' Description:       Returns the App object from VB
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       29/03/2004-08:18:52
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get Application() As App
    Set Application = App
End Property

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       SaveProperties
' Description:       Save property set to persistent storage
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12/11/2004-21:39:32
'
' Parameters :       PropSetName (String)
'                    PropSet (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function SaveProperties(PropSetName As Variant, PropSet As Variant) As Boolean
    xScriptProperties.Save CStr(PropSetName), PropSet
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       LoadProperties
' Description:       Load property set from persistent storage
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12/11/2004-21:40:13
'
' Parameters :       PropSetName (String)
'                    PropSet (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function LoadProperties(PropSetName As Variant, PropSet As Variant) As Boolean
    xScriptProperties.Load CStr(PropSetName), PropSet
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetProperty
' Description:       Get an individual property from the script's propset
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12/11/2004-21:41:17
'
' Parameters :       PropName (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetProperty(PropName As Variant) As Variant
    Dim p As ScriptProperty
    On Error Resume Next
    Set p = xScriptProperties.Find(CStr(PropName))
    If p Is Nothing Then
        Set GetProperty = Nothing
        LogMessage False, True, "Getting non-existent property: " & CStr(PropName)
        Err.Raise vbObjectError + 5, "GetProperty", "Getting non-existent property: " & CStr(PropName)
    Else
        GetProperty = p.Value
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       SetProperty
' Description:       Set an individual property in the script's propset
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12/11/2004-21:42:18
'
' Parameters :       PropName (String)
'                    Value (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub SetProperty(PropName As Variant, Value As Variant)
    Dim p As ScriptProperty
    Set p = xScriptProperties.Find(CStr(PropName))
    If Not (p Is Nothing) Then
        p.Value = Value
    Else
        LogMessage False, True, "Setting non-existent property: " & CStr(PropName)
        Err.Raise vbObjectError + 5, "SetProperty", "Setting non-existent property: " & CStr(PropName)
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Class_Initialize
' Description:       Class constructor
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12/11/2004-22:11:59
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Class_Initialize()
    Set xScriptProperties = New ScriptPropertySet
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ShowConfigDialog
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       13/11/2004-00:15:06
'
' Parameters :       PropSetName (String)
'                    Props (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ShowConfigDialog(PropSetName As String, PropSet As Variant) As Boolean
    With frmScriptCfg
        .Props = PropSet
        Set .PropSet = xScriptProperties
        .PropSetName = PropSetName
        .Show vbModal
        ShowConfigDialog = .RetVal
    End With
    Unload frmScriptCfg
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ChooseFromList
' Description:       Allow choice from a list of values through a dialog
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       8/11/2005-21:43:21
'
' Parameters :       sList (String)
'                    sDefault (String)
'                    sTitle (String)
'                    sHelp (String)
'                    bFixedList (Boolean)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ChooseFromList(sList As Variant, sDefault As Variant, sTitle As Variant, sHelp As Variant, _
    bFixedList As Variant, Optional sNewList As Variant, Optional sDelim As Variant = "\", Optional sValidateProc As Variant = "") As String
    
    Dim f As New frmScriptDropdown
    f.DialogTitle = CStr(sTitle)
    f.HelpText = CStr(sHelp)
    f.ValueList = CStr(sList)
    f.Selection = CStr(sDefault)
    f.FixedList = CBool(bFixedList)
    f.ValidateProc = CStr(sValidateProc)
    f.ListDelimiter = CStr(sDelim)
    
    f.Show vbModal
    ChooseFromList = f.Selection
    If Len(f.Selection) > 0 And (Not IsMissing(sNewList)) Then
        sNewList = f.NewChoices
    End If
    Unload f
End Function

Public Function HTMLDecode(sIn As Variant) As String
    HTMLDecode = HTMLDecodeString(CStr(sIn))
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetHTMLDocument
' Description:       Return the entire file in the form of an HTMLDocument object
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       3/1/2006-23:01:24
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetHTMLDocument()
    Dim sHTML As String
    Dim xDoc As New MSHTML.HTMLDocument
    
    On Error GoTo baleout
    sHTML = EntireFileHTML()
    Set xHTML = New MSHTML.HTMLDocument

'    Set xHTML = xDoc.createDocumentFromUrl(Session.InputFile.FileName, vbNullString)
    xHTML.Write sHTML

    Set GetHTMLDocument = xHTML
    Exit Function
baleout:
    MyMsgBox "Error in GetHTMLDocument(): " & Err.Description, vbCritical + vbOKOnly
    Set xHTML = Nothing
    Err.Raise Err.Number, "GetHTMLDocument", Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Function IsHTML() As Boolean

End Function

Private Function GetElementN(body As Variant, sTag As String, iIndex As Long) As Object
    Dim aElement As Object
    Dim cElements As Object 'IHTMLElementCollection
    Set cElements = body.All
    Dim i As Long, iFound As Long
    Dim oElement As Object
    If Not (cElements Is Nothing) Then
        iFound = 0
        For i = 0 To cElements.Length - 1
            Set oElement = cElements(i)
            If oElement Is Nothing Then
                MyMsgBox "element missing"
            End If
            If UCase$(oElement.tagName) = sTag Then
                iFound = iFound + 1
                If iFound = iIndex Then
                    Set GetElementN = oElement
                    Exit Function
                End If
            End If
        Next
    End If
    Set GetElementN = Nothing
    Exit Function
    
    Set cElements = cElements.getElementsByTagName(sTag)
    Set aElement = Nothing
    If cElements Is Nothing Then
        MyMsgBox "No element " & sTag & " found"
    Else
'        MyMsgBox "array of " & cElements.length & " rows"
        If iIndex <= cElements.Length Then
            Set aElement = cElements(iIndex - 1)
        End If
    End If
    Set GetElementN = aElement
End Function

Public Function HTMLForm(oBody As Variant, i As Long) As Object
    Set HTMLForm = GetElementN(oBody, "FORM", i)
End Function

Public Function HTMLTable(oBody As Variant, i As Long) As Object
    Set HTMLTable = GetElementN(oBody, "TABLE", i)
End Function

Public Function HTMLRow(oBody As Variant, i As Long) As Object
    Set HTMLRow = GetElementN(oBody, "TR", i)
End Function

Public Function HTMLCell(oBody As Variant, i As Long) As Object
    Set HTMLCell = GetElementN(oBody, "TD", i)
End Function

' atom ::= tag[(subscript)]
' expr = atom | expr.atom

Public Function HTMLSelect(oBody As Variant, sPath As String) As Object
    Dim a As Variant
    Dim oRet As Object: Set oRet = Nothing
    Dim i As Long
    Dim sTag As String
    Dim iNumSubscript As Long
    Dim sStringSubscript As String
    Dim sTmp As String
    Dim iParen As Long
    Dim oNodes As Object
    Dim oNode As Object
    Dim iCount As Long
    Dim bFound As Boolean
    
    Set HTMLSelect = Nothing
    On Error GoTo baleout
    Set oNodes = oBody
    a = Split(sPath, ".")
    For i = LBound(a) To UBound(a)
        iNumSubscript = 1
        sStringSubscript = ""
        sTag = a(i)
        iParen = InStr(sTag, "(")
        If iParen > 0 Then
            sTmp = Mid$(sTag, iParen + 1, Len(sTag) - iParen - 1)
            sTag = UCase$(Left$(sTag, iParen - 1))
            If IsNumeric(sTmp) Then
                iNumSubscript = Val(sTmp)
            Else
                sStringSubscript = sTmp
                If Left$(sStringSubscript, 1) = """" Then
                    sStringSubscript = Mid$(sStringSubscript, 2, Len(sStringSubscript) - 2)
                End If
            End If
        End If
        iCount = 0
        bFound = False
        Set oNodes = oNodes.childNodes   ' get direct descendants
        For Each oNode In oNodes
            If oNode.tagName = sTag Then
                If Len(sStringSubscript) = 0 Or sStringSubscript = oNode.id Then
                    iCount = iCount + 1
                    If iCount >= iNumSubscript Then
                        bFound = True
                        Exit For
                    End If
                End If
            End If
        Next
        If bFound Then
            Set oNodes = oNode
        Else
            Exit For
        End If
    Next
    If bFound Then
        Set HTMLSelect = oNode
    End If
    Exit Function
baleout:
    LogMessage True, True, "Error parsing HTML: " & Err.Description & vbCrLf & "Expression: " & sPath
    Abort
    Set HTMLSelect = Nothing
End Function

Public Function HTMLRowFields(oRow As Variant) As Variant
    Dim sTmp As String
' oRow should be a TR element
    If oRow.tagName <> "TR" Then
        LogMessage True, True, "HTMLRowFields: must be TR (got: " & oRow.tagName & ")"
        Abort
        Exit Function
    End If
    Dim a() As Variant
    ReDim a(0)
    Dim oCell As Object
    Dim oCells As Object
    Set oCells = oRow.childNodes
    For Each oCell In oCells
        If oCell.tagName = "TD" Or oCell.tagName = "TH" Then
            sTmp = oCell.innerText
            ReDim Preserve a(UBound(a) + 1)
            a(UBound(a)) = sTmp
        End If
    Next
    HTMLRowFields = a
End Function

Public Function LocalTZOffset() As Variant
    LocalTZOffset = GetCurrentTimeBias()    ' in minutes!
End Function

Public Function AppendScript(sFile As String) As Boolean
    Dim sModule As String
    On Error GoTo bad_err
    sModule = ActiveScriptModule()
    AppendScript = ScriptExtend(sModule, sFile)
goback:
    Exit Function
bad_err:
    AppendScript = False
    Resume goback
End Function

Public Function PromptForInput(sPrompt As Variant, sTitle As Variant, sText As Variant) As String
    Dim sTmp
    Dim bProgressOnTop As Boolean
    
    bProgressOnTop = SetProgressTopmost(False)
    On Error GoTo Cancelled

    sTmp = InputBox(CStr(sPrompt), CStr(sTitle), CStr(sText))
Cancelled:
    SetProgressTopmost bProgressOnTop
    PromptForInput = sTmp
End Function

Public Function HashStart() As Boolean
    HashStart = MD5HashStart()
End Function

Public Function HashAddString(sString As Variant) As Boolean
    HashAddString = MD5HashDataString(CStr(sString))
End Function

Public Function HashEnd() As String
    Dim bOut() As Byte
    Dim sHash As String
    If MD5HashEnd(bOut) Then
        sHash = HashHexString(bOut)
    Else
        sHash = ""
    End If
    HashEnd = sHash
End Function
