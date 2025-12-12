Attribute VB_Name = "MyScripting"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : MyScripting
'    Project    : MT2OFX
'
'    Description: Scripting Support Functions
'
'    Modified   : $Author: Colin $ $Date: 27/03/11 23:33 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/Scripting.bas 27    27/03/11 23:33 Colin $"
' $History: Scripting.bas $
' 
' *****************  Version 27  *****************
' User: Colin        Date: 27/03/11   Time: 23:33
' Updated in $/MT2OFX
' multiple account support
'
' *****************  Version 26  *****************
' User: Colin        Date: 15/11/10   Time: 0:29
' Updated in $/MT2OFX
'
' *****************  Version 25  *****************
' User: Colin        Date: 24/11/09   Time: 22:05
' Updated in $/MT2OFX
' for 3.6 beta
'
' *****************  Version 24  *****************
' User: Colin        Date: 15/06/09   Time: 19:25
' Updated in $/MT2OFX
' For transfer to new laptop
'
' *****************  Version 23  *****************
' User: Colin        Date: 25/11/08   Time: 22:24
' Updated in $/MT2OFX
' moving vss server!
'
' *****************  Version 21  *****************
' User: Colin        Date: 20/04/08   Time: 10:06
' Updated in $/MT2OFX
' For 3.5 beta 1
'
' *****************  Version 20  *****************
' User: Colin        Date: 7/12/06    Time: 15:07
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 18  *****************
' User: Colin        Date: 1/03/06    Time: 23:12
' Updated in $/MT2OFX
'
' *****************  Version 16  *****************
' User: Colin        Date: 2/11/05    Time: 23:03
' Updated in $/MT2OFX
' V3.4 beta 1
'
' *****************  Version 15  *****************
' User: Colin        Date: 8/05/05    Time: 12:43
' Updated in $/MT2OFX
' V3.3.8
'
' *****************  Version 14  *****************
' User: Colin        Date: 6/05/05    Time: 23:05
' Updated in $/MT2OFX
'</CSCC>

Const ProcInitialise As String = "Initialise"
Const ProcStartSession As String = "StartSession"
Const ProcEndSession As String = "EndSession"
Const ProcProcessStatement As String = "ProcessStatement"
Const ProcProcessTransaction As String = "ProcessTransaction"

Const ProcRecogniseTextFile As String = "RecogniseTextFile"
Const ProcLoadTextFile As String = "LoadTextFile"
Const ProcDescriptiveName As String = "DescriptiveName"
Const ProcConfigure As String = "Configure"
Const ProcValidationMessage As String = "ValidationMessage"

Const ProcOutputConfig As String = "ScriptConfig"
Const ProcCustomOutput As String = "CustomOutput"

Const GlobalModuleFile As String = "MT2OFX.vbs"
Const GlobalModuleFileJS As String = "MT2OFX.js"

Private xActiveModule As String
Private xActiveProc As String

Private sConfigMod As String

Private so As ScriptControl
Private sovb As ScriptControl
Private sojs As ScriptControl
Private sx As New ScriptEnv
Private GlobalLoaded As Boolean

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ScriptInit
' Description:       Initialise the scripting engine and load a module
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       9/9/2006-00:15:40
'
' Parameters :       sModule (String)
'                    sScript (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ScriptInit(sModule As String, sScript As String) As Boolean
    Dim iShowHelp As Long
    Dim sTmp As String
    Dim sScriptPath As String
    ScriptInit = False
    Dim m As Module
    Set sovb = frmMain.ScriptControl1
    Set sojs = frmMain.ScriptControl2
    On Error Resume Next
    sojs.AllowUI = True
    sojs.Language = "JavaScript"
    sovb.AllowUI = True
    sovb.Language = "VBScript"
' test of javascript
    If Cfg.EnableJavaScript Then
        If Right$(sScript, 3) = ".js" Then
            Set so = sojs
        Else
            Set so = sovb
        End If
    Else
        Set so = sovb
    End If
    so.Timeout = Cfg.ScriptTimeout
    Dim sOldActiveModule As String: sOldActiveModule = xActiveModule: xActiveModule = sModule
    Dim sOldActiveProc As String: sOldActiveModule = xActiveProc: xActiveProc = ""
    If sModule = GlobalModule Then
        so.Reset
        Err.Clear
        Set m = so.Modules(GlobalModule)
        If Err <> 0 Then
            With Err
                If Len(.HelpFile) > 0 Then
                    iShowHelp = vbMsgBoxHelpButton
                End If
                MyMsgBox "Unable to get GlobalModule " & sModule & ": " _
                    & .Description, vbCritical + vbOKCancel + iShowHelp, _
                    .Source, .HelpFile, .HelpContext
                xActiveModule = sOldActiveModule
                xActiveProc = sOldActiveProc
                Exit Function
            End With
        End If
        so.AddObject "Env", sx, "True"
    Else
        Set m = so.Modules(sModule)
        If Err = 0 Then ' script already loaded!
            ScriptInit = True
            xActiveModule = sOldActiveModule
            xActiveProc = sOldActiveProc
            Exit Function
        Else
            Err.Clear
            Set m = so.Modules.Add(sModule)
            If Err <> 0 Then
                With Err
                    If Len(.HelpFile) > 0 Then
                        iShowHelp = vbMsgBoxHelpButton
                    End If
                    MyMsgBox "Unable to add module " & sModule & ": " _
                        & .Description, vbCritical + vbOKCancel + iShowHelp, _
                        .Source, .HelpFile, .HelpContext
                    xActiveModule = sOldActiveModule
                    xActiveProc = sOldActiveProc
                    Exit Function
                End With
            End If
        End If
        Err.Clear
    End If
    Dim sGlobal As String
    Dim sf As New ScriptFile
'    On Error GoTo bad_open
'    Err.Clear
    sGlobal = sf.GetContents(sScript)
    
    On Error Resume Next
    Err.Clear
    m.AddCode sGlobal
    If Err <> 0 Then
        sTmp = "Error 0x" & Hex(Err.Number) & " from scripting module AddCode: " & Err.Description & vbCrLf _
            & "Source: " & Err.Source & vbCrLf _
            & "Last DLL error: 0x" & Hex(Err.LastDllError)
        LogMessage True, True, sTmp
    
        ReportScriptError "Loading Script", sModule
        With so.Error
            sTmp = "Script Syntax Error : " & .Number _
        & ": " & .Description & vbCrLf _
        & "at line " & .Line & " column " & .Column & vbCrLf _
        & "of file " & sScriptPath & ": " & vbCrLf _
        & "Text: " & .Text
            Debug.Print sTmp
            If Len(.HelpFile) > 0 Then
                iShowHelp = vbMsgBoxHelpButton
            End If
            sx.AbortRequested = (MyMsgBox(sTmp, vbCritical + vbOKCancel + iShowHelp, _
                .Source, .HelpFile, .HelpContext) = vbCancel)
            Err.Clear
            so.Error.Clear
        End With
    Else
        Debug.Print "Loaded " & sScript & " without errors."
        ScriptInit = CheckEps(sModule)
        If ScriptInit Then
            ScriptInit = ScriptInitialise(sModule)
        End If
    End If
go_back:
    xActiveModule = sOldActiveModule
    xActiveProc = sOldActiveProc
    Exit Function
bad_open:
    LogMessage True, False, _
        "Unable to open script file " & sScript & ": " & Err.Description, _
        "MT2OFX Script Error"
    Resume go_back
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ScriptInitialise
' Description:       Initialise a scripting module
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       9/9/2006-00:16:08
'
' Parameters :       sModule (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ScriptInitialise(sModule As String) As Boolean
    Set sx.Txn = Nothing
    Set sx.Statement = Nothing
    ScriptInitialise = CallScriptFunction(sModule, ProcInitialise)
    If sx.AbortRequested Then
        ScriptInitialise = False
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ScriptCustomOutput
' Description:       Call CustomOutput function
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       08/09/2010-23:53:21
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Function ScriptCustomOutput(sModule As String, Optional vArg) As Boolean
    Set sx.Txn = Nothing
    Set sx.Statement = Nothing
    ScriptCustomOutput = CallScriptFunction(sModule, ProcCustomOutput, vArg)
    If sx.AbortRequested Then
        ScriptCustomOutput = False
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ScriptExtend
' Description:       Append a script file to an existing module
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       9/9/2006-00:18:06
'
' Parameters :       sModule (String)
'                    sScript (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ScriptExtend(sModule As String, sScript As String) As Boolean
    Dim sCode As String
    Dim sScriptPath As String
    Dim iShowHelp As Long
    Dim iFile As Integer
    Dim m As Module
    Dim sTmp As String
    
' find the scripting control's module
    Set m = so.Modules(sModule)
    If Err <> 0 Then ' module not found
        ScriptExtend = False
        Exit Function
    End If

' get the new code
    Dim sf As New ScriptFile
'    On Error GoTo bad_open
    sCode = sf.GetContents(sScript)
    
' add the new code to the end of the existing module
    On Error Resume Next
    Err.Clear
    m.AddCode sCode

    If Err <> 0 Then
        ReportScriptError "Loading Script", sModule
        With so.Error
            sTmp = "Script Syntax Error : " & .Number _
        & ": " & .Description & vbCrLf _
        & "at line " & .Line & " column " & .Column & vbCrLf _
        & "of file " & sScriptPath & ": " & vbCrLf _
        & "Text: " & .Text
            Debug.Print sTmp
            If Len(.HelpFile) > 0 Then
                iShowHelp = vbMsgBoxHelpButton
            End If
            sx.AbortRequested = (MyMsgBox(sTmp, vbCritical + vbOKCancel + iShowHelp, _
                .Source, .HelpFile, .HelpContext) = vbCancel)
            Err.Clear
            so.Error.Clear
        End With
    Else
        Debug.Print "Appended " & sScript & " without errors."
        ScriptExtend = CheckEps(sModule)
    End If
    Exit Function
bad_open:
    LogMessage True, False, _
        "Unable to open script file " & sScript & ": " & Err.Description, _
        "MT2OFX Script Error"
End Function

Public Function ScriptStartSession(sModule As String) As Boolean
    Set sx.Txn = Nothing
    Set sx.Statement = Nothing
    ScriptStartSession = CallScriptFunction(sModule, ProcStartSession)
    If sx.AbortRequested Then
        ScriptStartSession = False
    End If
End Function

Public Function ScriptProcessStatement(sModule As String, s As Statement) As Boolean
    Set sx.Statement = s
    Set sx.Txn = Nothing
    ScriptProcessStatement = CallScriptFunction(sModule, ProcProcessStatement, s)
    If sx.AbortRequested Then
        ScriptProcessStatement = False
    End If
End Function

Public Function ScriptProcessTxn(sModule As String, t As Txn) As Boolean
    Set sx.Statement = t.Statement
    Set sx.Txn = t
    ScriptProcessTxn = CallScriptFunction(sModule, ProcProcessTransaction, t)
    If sx.AbortRequested Then
        ScriptProcessTxn = False
    End If
End Function

Public Function ScriptEndSession(sModule As String) As Boolean
    Set sx.Txn = Nothing
    Set sx.Statement = Nothing
    ScriptEndSession = CallScriptFunction(sModule, ProcEndSession)
    If sx.AbortRequested Then
        ScriptEndSession = False
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ScriptDescriptiveName
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       24/11/2003-23:46:00
'
' Parameters :       sModule (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ScriptDescriptiveName(sModule As String) As String
    Set sx.Txn = Nothing
    Set sx.Statement = Nothing
    Dim sTmp As String
    If CallScriptFunctionRet(sModule, ProcDescriptiveName, sTmp) Then
        ScriptDescriptiveName = sTmp
    Else
        ScriptDescriptiveName = ""
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ScriptProcessTextFile
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       06/12/2003-22:58:55
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ScriptProcessTextFile(sModule As String) As Boolean
    Dim vRet As Variant
    ScriptProcessTextFile = CallScriptFunctionRet(sModule, ProcLoadTextFile, vRet)
    ScriptProcessTextFile = ScriptProcessTextFile And CBool(vRet)
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ScriptCallConfig
' Description:       Call "Configure" in script
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       13/11/2004-00:35:45
'
' Parameters :       sModule (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ScriptCallConfig(sModule As String) As Boolean
    sConfigMod = sModule
    ScriptCallConfig = CallScriptFunction(sModule, ProcConfigure)
    sConfigMod = ""
End Function

Public Function ScriptDoValidate(sModule As String, sProc As String, vArg As Variant) As Boolean
    On Error GoTo bad
    Dim vRet As Variant
    ScriptDoValidate = CallScriptFunction(sModule, sProc, vArg, vRet)
    If ScriptDoValidate Then
        ScriptDoValidate = CBool(vRet)
    End If
    Exit Function
bad:
    ScriptDoValidate = False
End Function

Public Function ScriptValidationMessage(sModule As String) As String
    Dim vRet As Variant
    On Error Resume Next
    ScriptValidationMessage = ""
    If CallScriptFunctionRet(sModule, ProcValidationMessage, vRet) Then
        ScriptValidationMessage = CStr(vRet)
    End If
End Function
Public Function FindEp(sModule As String, sProc As String) As Procedure
    Dim p As Procedure
    Set FindEp = Nothing
    For Each p In so.Modules(sModule).Procedures
        If p.Name = sProc Then
            Set FindEp = p
            Exit Function
        End If
    Next
End Function
Public Function EpDefined(sModule As String, sProc As String) As Boolean
    Dim p As Procedure
    Set p = FindEp(sModule, sProc)
    EpDefined = Not (p Is Nothing)
End Function
Private Function CheckEps(sModule As String) As Boolean
    Dim bErr As Boolean
    bErr = False
    Dim p As Procedure
    Set p = FindEp(sModule, ProcInitialise)
    If Not (p Is Nothing) Then
        If p.NumArgs <> 0 Then
            MyMsgBox GetString(119, sModule, ProcInitialise, 0)
            bErr = True
        End If
    End If
    Set p = FindEp(sModule, ProcStartSession)
    If Not (p Is Nothing) Then
        If p.NumArgs <> 0 Then
            MyMsgBox GetString(119, sModule, ProcStartSession, 0)
            bErr = True
        End If
    End If
    Set p = FindEp(sModule, ProcProcessStatement)
    If Not (p Is Nothing) Then
        If p.NumArgs <> 1 Then
            MyMsgBox GetString(119, sModule, ProcProcessStatement, 1)
            bErr = True
        End If
    End If
    Set p = FindEp(sModule, ProcProcessTransaction)
    If Not (p Is Nothing) Then
        If p.NumArgs <> 1 Then
            MyMsgBox GetString(119, sModule, ProcProcessTransaction, 1)
            bErr = True
        End If
    End If
    Set p = FindEp(sModule, ProcEndSession)
    If Not (p Is Nothing) Then
        If p.NumArgs <> 0 Then
            MyMsgBox GetString(119, sModule, ProcEndSession, 0)
            bErr = True
        End If
    End If
    CheckEps = Not bErr
End Function
Public Function CallScriptFunction(sModule As String, sProc As String, Optional vArg As Variant, Optional vRet As Variant) As Boolean
    If Not EpDefined(sModule, sProc) Then
        CallScriptFunction = True
        Exit Function
    End If
    Dim sTmp As String
    Dim iShowHelp As Long
    CallScriptFunction = False
    Dim sOldActiveModule As String: sOldActiveModule = xActiveModule: xActiveModule = sModule
    Dim sOldActiveProc As String: sOldActiveProc = xActiveProc: xActiveProc = sProc
    On Error Resume Next
    Set sx.Cfg = Cfg
    Set sx.Bcfg = Bcfg
    Err.Clear
    so.Error.Clear
    Dim v As Variant
    If IsMissing(vArg) Then
        v = so.Modules(sModule).Run(sProc)
    Else
        v = so.Modules(sModule).Run(sProc, vArg)
    End If
    If Not IsMissing(vRet) Then vRet = v
    If Err = 0 Then
        CallScriptFunction = True
    Else
        If so.Error.Number = 0 Then
            If Len(Err.HelpFile) > 0 Then
                iShowHelp = vbMsgBoxHelpButton
            End If
            sx.AbortRequested = (MyMsgBox(Err.Description, vbOKCancel + vbCritical + iShowHelp, _
                Err.Source, Err.HelpFile, Err.HelpContext) = vbCancel)
            Debug.Print "Error from " & sProc & ": " & Err.Description
            sTmp = sx.LastLine
            Debug.Print "Last input line: " & sTmp
            Debug.Assert False
        Else
            ReportScriptError sProc, sModule
        End If
        so.Error.Clear
        Err.Clear
    End If
    xActiveModule = sOldActiveModule
    xActiveProc = sOldActiveProc
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CallScriptFunctionRet
' Description:       Calls a script function and returns its return value
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       30/11/2003-23:22:12
'
' Parameters :       sModule (String)
'                    sProc (String)
'                    vRet (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function CallScriptFunctionRet(sModule As String, sProc As String, vRet As Variant) As Boolean
    If Not EpDefined(sModule, sProc) Then
        CallScriptFunctionRet = False
        Exit Function
    End If
    Dim sTmp As String
    Dim iShowHelp As Long
    CallScriptFunctionRet = False
    Dim sOldActiveModule As String: sOldActiveModule = xActiveModule: xActiveModule = sModule
    On Error Resume Next
    Set sx.Cfg = Cfg
    Set sx.Bcfg = Bcfg
    Err.Clear
    so.Error.Clear
    vRet = so.Modules(sModule).Eval(sProc)
    If Err = 0 Then
        CallScriptFunctionRet = True
    Else
        If so.Error.Number = 0 Then
            If Len(Err.HelpFile) > 0 Then
                iShowHelp = vbMsgBoxHelpButton
            End If
            sx.AbortRequested = (MyMsgBox(Err.Description, vbOKCancel + vbCritical + iShowHelp, _
                Err.Source, Err.HelpFile, Err.HelpContext) = vbCancel)
            Debug.Print "Error from " & sProc & ": " & Err.Description
            sTmp = sx.LastLine
            Debug.Print "Last input line: " & sTmp
            Debug.Assert False
        Else
            ReportScriptError sProc, sModule
        End If
        so.Error.Clear
        Err.Clear
    End If
    xActiveModule = sOldActiveModule
End Function


Public Sub ReportScriptError(sSource As String, sModule As String)
    Dim sTmp As String
    Dim iShowHelp As Long
' 20050707 CS: Improved message: now includes err num in hex, plus the module name
    With so.Error
        sTmp = "Error " & .Number & " (0x" & Hex(.Number) & "): " & .Description
        sTmp = sTmp & vbCrLf & "At line " & .Line & " col " & .Column
        If Len(.Text) > 0 Then
            sTmp = sTmp & vbCrLf & .Text
        End If
        If Len(sModule) > 0 Then
            sTmp = sTmp & vbCrLf & "In module " & sModule
        End If
        If Len(sSource) > 0 Then
            sTmp = sTmp & vbCrLf & "Last function was " & sSource
        End If
        sTmp = sTmp & vbCrLf & "Last input line was " & sx.LastLine
        Debug.Print sTmp
        LogMessage False, True, sTmp, ""
        If Len(.HelpFile) > 0 Then
            iShowHelp = vbMsgBoxHelpButton
        End If
        sx.AbortRequested = (MyMsgBox(sTmp, vbOKCancel + vbCritical + iShowHelp, _
            .Source, .HelpFile, .HelpContext) = vbCancel)
        Debug.Assert False
    End With
End Sub

Public Function GetScriptEnv() As ScriptEnv
    Set GetScriptEnv = sx
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       InitialiseScripting
' Description:       Load global module if required
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       25/11/2003-22:27:48
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function InitialiseScripting() As Boolean
    If Not GlobalLoaded Then
        If ScriptInit(GlobalModule, Cfg.ScriptPath & "\" & GlobalModuleFile) Then
            GlobalLoaded = True
            If Cfg.EnableJavaScript Then
                If ScriptInit(GlobalModule, Cfg.ScriptPath & "\" & GlobalModuleFileJS) Then
                    Debug.Print "Error loading javascript library"
                End If
            End If
        Else
' message 118 is "Processing aborted" - we must be able to improve on that
            LogMessage True, True, GetString(118), AppName
        End If
    End If
    sx.AbortRequested = False
    InitialiseScripting = GlobalLoaded
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ResetScripting
' Description:       Reset scripting environment - forget loaded modules!
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       29/03/2004-22:45:47
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub ResetScripting()
    If Not (so Is Nothing) Then
        so.Reset
        GlobalLoaded = False
    End If
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ChooseScript
' Description:       choose a script file (*.vbs)
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       03/01/2004-21:17:25
'
' Parameters :       sScript (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ChooseScript(oCd As CommonDialog, ByVal sScript As String) As String
    Dim iLen As Long
    Dim sTmp As String
    With oCd
' 20050505 CS: Dialog strings localised
        .DialogTitle = LoadResStringL(550)
        .InitDir = Cfg.ScriptPath
        If sScript = "" Then
            .FileName = ""
        Else
            .FileName = FindScript(sScript)
        End If
' 20050505 CS: Dialog strings localised
        .Filter = LoadResStringL(551)
        .FilterIndex = 1
        .Flags = cdlOFNFileMustExist + cdlOFNLongNames + cdlOFNExplorer _
            + cdlOFNHideReadOnly + cdlOFNPathMustExist
        .ShowOpen
        If .FileName <> "" Then
            sTmp = Cfg.ScriptPath & "\"
            iLen = Len(sTmp)
            If UCase$(Left$(.FileName, iLen)) = UCase$(sTmp) Then
                sScript = Mid$(.FileName, iLen + 1)
            Else
                sScript = .FileName
            End If
        End If
    End With
    ChooseScript = sScript
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       FindScript
' Description:       Find a script, either as an absolute path or
'                    relative to the app directory
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       13/01/2004-13:47:01
'
' Parameters :       sScript (String)
' 20060124 CS: Now returns empty string if file cannot be found
'--------------------------------------------------------------------------------
'</CSCM>
Public Function FindScript(sScript As String) As String
    Dim sTmp As String
    If PathIsRelative(sScript) Then
        sTmp = Cfg.ScriptPath & "\" & sScript
    Else
        sTmp = sScript
    End If
    If Dir(sTmp, vbNormal) = "" Then
        FindScript = ""
    Else
        FindScript = sTmp
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ScriptShowAbout
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12-Jan-2005-20:47:13
'
' Parameters :       sScript (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ScriptShowAbout(sScript As String) As Boolean
    ScriptShowAbout = False
    Dim sAbout As String
    Dim sMod As String
    If sScript = "" Then
        Exit Function
    End If
    ResetScripting
    Dim sx As ScriptEnv
    Set sx = GetScriptEnv()
    sMod = ScriptModuleName(sScript)
    If Not InitialiseScripting() Then
        Exit Function
    End If
    If Not ScriptInit(sMod, sScript) Then
        Exit Function
    End If
    If EpDefined(sMod, "DescriptiveName") Then
        sAbout = ScriptDescriptiveName(sMod)
        MyMsgBox sAbout, , sMod
        ScriptShowAbout = True
    Else
        MyMsgBox LoadResStringL(502), , sMod
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ScriptShowSettings
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12-Jan-2005-20:47:40
'
' Parameters :       sScript (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ScriptShowSettings(sScript As String) As Boolean
    Dim sMod As String
    ScriptShowSettings = False
    If sScript = "" Then
        Exit Function
    End If
    ResetScripting
    Dim sx As ScriptEnv
    Set sx = GetScriptEnv()
    sx.AbortRequested = False
    sMod = ScriptModuleName(sScript)
    If Not InitialiseScripting() Then
        Exit Function
    End If
    If Not ScriptInit(sMod, sScript) Then
        Exit Function
    End If
    If EpDefined(sMod, "Configure") Then
        ScriptCallConfig sMod
        ScriptShowSettings = True
    Else
        MyMsgBox LoadResStringL(501), , sMod
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ActiveScriptModule
' Description:       Returns currently active module name
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       1/3/2007-20:58:46
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ActiveScriptModule() As String
    ActiveScriptModule = xActiveModule
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ActiveConfigModule
' Description:       [type_description_here]
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       18/03/2011-14:04:09
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ActiveConfigModule() As String
    ActiveConfigModule = sConfigMod
End Function
