Attribute VB_Name = "MainProg"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : MainProg
'    Project    : MT2OFX
'
'    Description: MT2OFX Entry Point
'
'    Modified   : $Author: Colin $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/Main.bas 23    15/11/10 0:11 Colin $"
' $History: Main.bas $
' 
' *****************  Version 23  *****************
' User: Colin        Date: 15/11/10   Time: 0:11
' Updated in $/MT2OFX
'
' *****************  Version 22  *****************
' User: Colin        Date: 5/01/10    Time: 17:56
' Updated in $/MT2OFX
' Fixed problem causing main form to restart when closing prog
'
' *****************  Version 21  *****************
' User: Colin        Date: 6/10/09    Time: 0:30
' Updated in $/MT2OFX
' new params for watcher
' exit codes for watcher
'
' *****************  Version 20  *****************
' User: Colin        Date: 15/06/09   Time: 19:25
' Updated in $/MT2OFX
' For transfer to new laptop
'
' *****************  Version 19  *****************
' User: Colin        Date: 25/11/08   Time: 22:23
' Updated in $/MT2OFX
' moving vss server!
'
' *****************  Version 17  *****************
' User: Colin        Date: 20/04/08   Time: 10:05
' Updated in $/MT2OFX
' For 3.5 beta 1
'
' *****************  Version 16  *****************
' User: Colin        Date: 7/12/06    Time: 15:07
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2

'</CSCC>

Private Const ModHeader As String = "$Header: /MT2OFX/Main.bas 23    15/11/10 0:11 Colin $"

Public Declare Function GetUserDefaultUILanguage Lib "kernel32" () As Integer
Public Declare Function GetSystemDefaultUILanguage Lib "kernel32" () As Integer
Public Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Public Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer

' mutex stuff to prevent multiple instances
Private Const MutexName As String = "MT2OFX-Mutex"
Private Declare Function CreateMutex Lib "kernel32" _
        Alias "CreateMutexA" _
       (ByVal lpMutexAttributes As Long, _
        ByVal bInitialOwner As Long, _
        ByVal lpName As String) As Long
'variable constant to match if the mutex exists
Private Const ERROR_ALREADY_EXISTS = 183&
Private hMutex As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long

Public Const MinScriptDllVersion As Long = &H50006  ' version 5.6

Public HTMLClipboardData As String
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Public argv() As String
Public argc As Long

' 20090917 CS: Hold command line parameters here
Public CmdParams As New Parameters
Dim sUsage As String

Public NODATE As Date

' AbortRequested is set to indicate we should break out of statement/
' transaction processing loop
'Public AbortRequested As Boolean

Sub MigrateOneFile(sFile As String, sFromDir As String, sToDir As String)
    Dim fso As New Scripting.FileSystemObject
    Dim sFrom As String
    Dim sTo As String
    On Error GoTo badmove
    sFrom = sFromDir & "\" & sFile
    If Not fso.FileExists(sFrom) Then
        Exit Sub
    End If
    sTo = sToDir & "\" & sFile
    If fso.FileExists(sTo) Then
        Exit Sub
    End If
    MyMsgBox "Migrating " & sFile & " from " & sFromDir & " to " & sToDir
    fso.MoveFile sFrom, sTo
    Exit Sub
badmove:
    MyMsgBox "Error migrating " & sFile & ": " & Err.Description
End Sub
Sub CheckMigrateDataFiles()
    Dim sFromDir As String, sToDir As String
    sFromDir = App.Path
    sToDir = GetSpecialFolder(CSIDL_APPDATA) & "\MT2OFX"
    Dim fso As New Scripting.FileSystemObject
    If Not fso.FolderExists(sToDir) Then
        fso.CreateFolder sToDir
    End If
    MigrateOneFile "mt2ofx.ini", sFromDir, sToDir
    MigrateOneFile "sprops.ini", sFromDir, sToDir
    MigrateOneFile "payeemap.dat", sFromDir, sToDir
    MigrateOneFile "support.txt", sFromDir, sToDir
    MigrateOneFile "log.txt", sFromDir, sToDir
' force reload of ini file!
    Set Cfg = New ProgramConfig
End Sub

Sub ConfigureUILanguage()
' set messages to use user's language (if poss!)
    Dim lTmp As Long
    Dim sResFile As String
    lTmp = GetConfigLang()
    If lTmp = 0 Then
' following line only for nt/2k/xp/me, hence we either check the os
' version, or we fall back to universal code, as we do...
'        SetProgLocale GetUserDefaultUILanguage()
        SetProgLocale GetUserDefaultLangID()
    Else
        SetProgLocale lTmp
    End If
    sResFile = GetMyString("General", "ResFile", "")
    If sResFile <> "" Then
        If PathIsRelative(sResFile) Then
            sResFile = Cfg.ResourcePath & "\" & sResFile
        End If
        If SetLanguageFile(sResFile) Then
' ok, we use the file
        Else
            SetLanguageFile ""
        End If
    End If
End Sub

Private Function ProcessCommandLine(argv() As String) As Boolean
    Dim i As Integer
    Dim s As String
    Dim bErr As Boolean
    Dim sTmp As String
    Dim sErrMsg As String
    
    Set CmdParams = New Parameters
    bErr = False
    i = LBound(argv)
' when running from the IDE the first argument is vb6.exe, the second is the project (.vbp) file
    If Right(argv(i), 8) = "\vb6.exe" Then
        i = i + 2
    Else    ' skip exe path argv[0]
        i = i + 1
    End If
    Do While i <= UBound(argv)
        s = argv(i)
        If Left$(s, 1) = "/" Then
            Select Case LCase$(s)
            Case "/clipboard"
                CmdParams.FromClipboard = True
            Case "/out"
                If i < UBound(argv) Then
                    CmdParams.OutputFile = argv(i + 1)
                    i = i + 1
                Else
                    sErrMsg = "Output file missing"
                    bErr = True
                End If
            Case "/type"
                If i < UBound(argv) Then
                    sTmp = UCase(argv(i + 1))
                    If sTmp = "OFX" Or sTmp = "QIF" Or sTmp = "QFX" Or sTmp = "OFC" Then
                        CmdParams.OutputType = sTmp
                    ElseIf IsValidCustomFormat(sTmp) Then
                        CmdParams.OutputType = sTmp
                    Else
                        sErrMsg = "Unknown output type '" & argv(i + 1) & "'"
                        bErr = True
                    End If
                    i = i + 1
                Else
                    sErrMsg = "Output type missing"
                    bErr = True
                End If
            Case "/script"
                If i < UBound(argv) Then
                    CmdParams.ScriptName = argv(i + 1)
                    i = i + 1
                Else
                    sErrMsg = "Script name missing"
                    bErr = True
                End If
            Case "/quiet"
                CmdParams.Quiet = True
            Case Else
                sErrMsg = "Unknown option '" & s & "'"
                bErr = True
            End Select
        Else
            If Len(CmdParams.InputFile) > 0 Then
                sErrMsg = "Multiple input files: prev=" & CmdParams.InputFile & ", new=" & s
                bErr = True
            Else
                CmdParams.InputFile = s
            End If
        End If
        If bErr Then
            Exit Do
        End If
        i = i + 1
    Loop
    If bErr Then
        sUsage = "Usage: mt2ofx.exe [/quiet] [/out ""outfile"" [/type QIF|OFC|OFX|QFX]] [""infile""|/clipboard]"
        MyMsgBox sUsage, vbCritical + vbOKOnly, "MT2OFX"
    End If
    ProcessCommandLine = Not bErr
End Function

' this is the entry point. If we are called with a parameter, we assume it
' is the name of an MT940 file to convert and do our work silently (well,
' fairly quietly anyway.) If we are started without parameters, the GUI is
' shown.

Sub Main()
    Dim sCmd As String
    On Error GoTo bigerror
    Dim sTmp As String
    Dim bClip As Boolean
    Dim iExit As Integer

    Call ConfigureUILanguage
'Place in startup code (Form_Load or Sub Main):

   ' Check if this is the first instance:
   If Not WeAreAlone(MutexName) Then
        MyMsgBox LoadResStringL(155), vbCritical + vbOKOnly
        CleanUpMutex
        Terminate 1
   End If
'    hMutex = CreateMutex(0&, 1&, MutexName)
'    If Err.LastDllError = ERROR_ALREADY_EXISTS Then
'        MyMsgBox LoadResStringL(155), vbCritical + vbOKOnly
'        CleanUpMutex
'        End
'    End If
    
    NODATE = DateSerial(0, 0, 0)
        
' 20061129 CS: migrate data files to appdata
    CheckMigrateDataFiles

    If Not Cfg.Load() Then
        MyMsgBox LoadResStringL(108), vbCritical + vbOKOnly
        CleanUpMutex
        Terminate 2
    End If

    iExit = 0
' get command line - ignore if in IDE
    sCmd = GetCommandLine()
    DBCSLog sCmd, "Command Line"
    argv = ParseCommandLine(sCmd)
    argc = UBound(argv) + 1
    If argc > 1 Then
        LogMessage False, True, sCmd, "Command Line"
    End If
    If InStr(argv(0), "\vb6.exe") > 0 Then
        argc = 1
    End If
    
    If Not Cfg.NoScriptingVersionCheck Then
        If Not CheckScriptingVersion("vbscript.dll", MinScriptDllVersion) Then
            CleanUpMutex
            Terminate 3
        End If
    End If

' 20070208 CS: maybe run QuickStart now?
    If Cfg.ShowQuickStart Then
        frmQuickStart.Show vbModal
        Cfg.ShowQuickStart = False
        Cfg.Save
    End If

    Dim sIn As String
    Dim sOut As String
    Dim sType As String, iType As Long

    If Not ProcessCommandLine(argv) Then
        iExit = 9
        GoTo baleout
    End If

' if no input file then run GUI mode
    If Len(CmdParams.InputFile) = 0 And Not (CmdParams.FromClipboard) Then
' modal forms are never shown in the task bar so we can't use vbModal!
        frmMain.Show ' vbModal
        Do While frmMain.Visible
            Sleep 100
            DoEvents
            If Not IsFormLoaded("frmMain") Then
                Exit Do
            End If
        Loop
        Terminate 0
    End If

' what type of file are we making? If specified on the command line, use that, otherwise use the default from the config
    sType = CmdParams.OutputType
    If Len(sType) = 0 Then
        sType = Cfg.OutputFileType
    End If

' get the output file name. if a directory is specified then the filename will be copied from the input file,
' with the extension changed appropriately for the output file tye
    sOut = CmdParams.OutputFile
    If Len(sOut) > 0 Then
        If IsDirectory(sOut) Then
            If Right(sOut, 1) <> "\" Then sOut = sOut & "\"
            sOut = sOut + ChangeExtension(GetFilename(sIn), "OFX")
        End If
    End If
    
    If CmdParams.FromClipboard Then
        sIn = CaptureTextClipboard()
        DBCSLog sIn, "Clipboard temp input"
        If sIn = "" Then
            iExit = 4
            GoTo baleout
        End If
        HTMLClipboardData = GetHTMLClipboard(True)
        If Len(sOut) = 0 Then
            sOut = GetOutputFile("", 0, sType, iType)
        End If
    Else
        DBCSLog CmdParams.InputFile, "Input File"
        sIn = ShortFileNameToLong(CmdParams.InputFile)
        DBCSLog sIn, "Input Long name"
        If Len(sOut) = 0 Then
            sOut = GetOutputFile(sIn, 0, sType, iType)
        End If
    End If

    DBCSLog sOut, "Output file name"
    If sOut = "" Then
        iExit = 5
        GoTo baleout
    End If
    
    Debug.Print "In: " & sIn & vbCrLf & "Out: " & sOut
    
    If Process(sIn, sOut, iType, sType) Then
        DoImport sOut, Cfg.NoConfirmImport
' make sure we give the import prog enough time to get started
        If Cfg.PromptForOutput = OUTFILE_TEMP Then
            DoEvents
            Sleep Cfg.TempFileDelay
        End If
    Else
        iExit = 6
    End If
baleout:
    If Cfg.PromptForOutput = OUTFILE_TEMP Then
        RemoveTempFile sOut
    End If
    If CmdParams.FromClipboard Then
        RemoveTempFile sIn
    End If
    CloseDBCSLog
    CloseLogFile
    CleanUpMutex
    Terminate iExit
    End
bigerror:
    MyMsgBox Err.Description, vbOKOnly + vbCritical, "Big MT2OFX Problem"
    Resume baleout
End Sub

Private Function IsFormLoaded(sForm As String) As Boolean
    Dim oFrm As Form

    IsFormLoaded = False
    For Each oFrm In Forms
        If oFrm.Name = sForm Then
            IsFormLoaded = True
            Exit For
        End If
    Next oFrm
End Function

Public Sub CleanUpMutex()
    If hMutex <> 0 Then
        ReleaseMutex hMutex
        CloseHandle hMutex
        hMutex = 0
    End If
End Sub

Private Function WeAreAlone(ByVal sMutex As String) As Boolean
   ' Don't call Mutex when in VBIDE because it will apply
   ' for the entire VB IDE session, not just the app's
   ' session.
   If InDevelopment Then
      WeAreAlone = Not (App.PrevInstance)
   Else
      ' Ensures we don't run a second instance even
      ' if the first instance is in the start-up phase
      hMutex = CreateMutex(ByVal 0&, 1, sMutex)
      If (Err.LastDllError = ERROR_ALREADY_EXISTS) Then
         CleanUpMutex
      Else
         WeAreAlone = True
      End If
   End If
End Function

Private Sub Terminate(iExit As Integer)
    If InDevelopment Then
        If iExit <> 0 Then
            MsgBox "Exit code: " & CStr(iExit)
        End If
        End
    Else
        ExitProcess iExit
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       IsValidCustomFormat
' Description:
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       08/05/2010-15:52:06
'
' Parameters :       sName (String)
'--------------------------------------------------------------------------------
'</CSCM>
Private Function IsValidCustomFormat(sName As String) As Boolean
    Dim xCust As New CustomOutputList
    xCust.Load (Cfg.CustomOutputPath)
    IsValidCustomFormat = Not (xCust.FindByFormatName(sName) Is Nothing)
End Function

