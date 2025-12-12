'
' script to transfer accounting hub file by sftp using winscp
' Colin Smale, IBM Global Services, 5 Nov 2009
'
' This script must be installed in the same directory as the GeneralLedgerProcessor executable and DLLs as it
' uses the same configuration file: GeneralLedgerProcessor.exe.config which can be edited via the XML Configurator.
'
Option Explicit
Const exe = "C:\Program Files\WinSCP\winscp.com"
Const config = "GeneralLedgerProcessor.exe.config"
Const log4net_setting = "/configuration/appSettings/add[@key='Log4NetFile']"
Const config_section = "/configuration/castle/components/component[@id='gl.configurationsettings']/parameters"
Const logconfig = "log4net.xml"
Const log4net_filesetting = "/log4net/appender[@type='log4net.Appender.RollingFileAppender']/param[@name='File']"
Const logfile = "AH_Transfer.log"

Dim sTmp: sTmp = WScript.ScriptFullName
Dim iTmp: iTmp = InStrRev(sTmp, "\")
Dim base: base = Left(sTmp, iTmp)

Dim fso: Set fso=CreateObject("Scripting.FileSystemObject")
Dim re: Set re=New RegExp
Dim wsh: Set wsh = CreateObject("WScript.Shell")
Dim wshEnv: Set wshEnv = wsh.Environment("Process")
Dim tempdir
tempdir = wshEnv("TEMP")
If Len(tempdir) = 0 Then
    tempdir = wshEnv("TMP")
End If

' The WinSCP log file and ini file are created and deleted again by this script
Dim AHlogfile: AHlogfile = tempdir & "AHlog.xml"
Dim AHinifile: AHinifile = tempdir & "WinSCP.ini"

' The error log file is appended to by this script
Dim AHerrlog
AHerrlog = base & logfile  ' temp in case of errors before real file is found in GetConfig

' config
Dim AHremotedir, AHremotehost, AHsslkey, AHlocaldir, AHdonedir, AHpat
If Not GetConfig(base & config) Then
    AppendToLog "ERROR: Configuration errors, cannot continue."
    WScript.Quit 2
End If

' destination
'Dim AHremotedir: AHremotedir="from_xilion"
'Dim AHremotehost: AHremotehost="scp://drm_ftp:gFqSV9aM@10.68.14.97/"
'Dim AHsslkey: AHsslkey = "ssh-rsa 1024 8a:42:af:95:95:82:9d:8d:77:71:cd:b0:d4:36:19:f1"
' AHremotehost="sftp://colin:colin@localhost/"

Dim fsoLocalDir, fsoFile
Dim nFiles: nFiles = 0
Dim bErr: bErr = False

' kill old log file
If fso.FileExists(AHlogfile) Then
    fso.DeleteFile AHlogfile, True
End If

' kill ini file to clear the cached ssl fingerprints
If fso.FileExists(AHinifile) Then
    fso.DeleteFile AHinifile, True
End If

' get input directory
If Not fso.FolderExists(AHlocaldir) Then
    fso.CreateFolder(AHlocaldir)
End If
Set fsoLocalDir = fso.GetFolder(AHlocaldir)

' loop through available files
re.Pattern = AHpat
re.IgnoreCase = True
For Each fsoFile in fsoLocalDir.Files
    If re.Test(fsoFile.Name) Then
        If Not ProcessFile(fsoFile) Then
            bErr = True
            Exit For
        End If
        nFiles = nFiles + 1
    End If
Next

' kill ini file to prevent the cached ssl fingerprints hanging around
If fso.FileExists(AHinifile) Then
    fso.DeleteFile AHinifile, True
End If

' all done, hopefully with no errors
If nFiles > 0 Then
    AppendToLog nFiles & " files transferred."
End If

If bErr Then
    WScript.Quit 1
Else
    WScript.Quit 0
End If

' function to process a single file
Function ProcessFile(fsoFile)
    If TransferFile(fsoFile.Path) Then
        ProcessFile = MoveToDone(fsoFile)
    Else
        ProcessFile = False
    End If
End Function

' function to transfer a single file
Function TransferFile(sFile)
    Dim sCmd, oExec, sExt

    sCmd = QuoteString(exe) & " /log=" & QuoteString(AHlogfile) & " /console /ini=" & QuoteString(AHinifile)
    sExt = GetExtension(sFile)
    Set oExec = wsh.Exec(sCmd)
    oExec.StdIn.Write "option batch abort" & vbLf
    If Len(AHsslkey) > 0 Then
        oExec.StdIn.Write "open " & AHremotehost & " -hostkey=" & QuoteString(AHsslkey) & vbLf
    Else
        oExec.StdIn.Write "open " & AHremotehost & vbLf
    End If
    If Len(AHremotedir) > 0 Then
        oExec.StdIn.Write "cd " & AHremotedir & vbLf
    End If
    oExec.StdIn.Write "option transfer binary" & vbLf
    oExec.StdIn.Write "put " & sFile & " *.tmp -preservetime -resume" & vbLf
    oExec.StdIn.Write "mv *.tmp *." & sExt & vbLf
    oExec.StdIn.Write "close" & vbLf
    oExec.StdIn.Write "exit" & vbLf

    Dim sOutput: sOutput = oExec.StdOut.ReadAll

    Do While oExec.Status = 0
        WScript.Sleep 100
    Loop

    If oExec.ExitCode = 0 Then
        AppendToLog "Successfully transferred file " & sFile
    Else
        AppendToLog "Error transferring file " & sFile
        AppendToLog sOutput
    End If
    TransferFile = (oExec.ExitCode = 0)
End Function

Function QuoteString(s)
    QuoteString = """" & s & """"
End Function

Function GetExtension(sFile)
    Dim iDot
    iDot = InStrRev(sFile, ".")
    If iDot > 0 Then
        GetExtension = Mid(sFile, iDot+1)
    Else
        GetExtension = ""
    End If
End Function

' function to move a file to the "done" directory
Function MoveToDone(fsoFile)
    If Not fso.FolderExists(AHdonedir) Then
        fso.CreateFolder(AHdonedir)
    End If
    Dim sDest
    sDest = AHdonedir & "\" & fsoFile.Name
    If fso.FileExists(sDest) Then
        fso.DeleteFile sDest, True
    End If
    fso.MoveFile fsoFile.Path, sDest
    MoveToDone = True
End Function

Function AppendToLog(sText)
    Dim f
    Set f = fso.OpenTextFile(AHerrlog, 8, True)
    f.Write Now()
    f.Write " "
    f.WriteLine sText
    f.Close
End Function

Function GetConfig(sConfig)
    GetConfig = False
    Dim xmlConfig: Set xmlConfig = CreateObject("MSXML2.DOMDocument")
' load config file
    xmlConfig.async = False
    If Not xmlConfig.load(sConfig) Then
        AppendToLog "ERROR: unable to load config file " & sConfig & ": error=" & xmlConfig.parseError.reason
        Exit Function
    End If

' get log4net config file
    Dim xNode, sLog
    Set xNode = xmlConfig.selectSingleNode(log4net_setting)
    If xNode Is Nothing Then
        sLog = base & logconfig
    Else
        sLog = xNode.getAttribute("value")
    End If
    AHerrlog = GetLogFile(sLog, logfile)

' get our config section
    Set xNode = xmlConfig.selectSingleNode(config_section)
    If xNode Is Nothing Then
        AppendToLog "ERROR: can't find config node for " & config_section
        Exit Function
    End If

' read our config parameters
    Dim sHost: sHost = NodeText(xNode.selectSingleNode("ftphost"))
    Dim sHostKey: sHostKey = NodeText(xNode.selectSingleNode("hostkey"))
    Dim sLogin: sLogin = NodeText(xNode.selectSingleNode("login"))
    Dim sPassword: sPassword = NodeText(xNode.selectSingleNode("password"))
    Dim sPattern: sPattern = NodeText(xNode.selectSingleNode("pattern"))
    Dim sRemoteDir: sRemoteDir = NodeText(xNode.selectSingleNode("remotedir"))
    Dim sInputDir: sInputDir = NodeText(xNode.selectSingleNode("TargetFilePath"))   ' shared with main program
    Dim sDoneDir: sDoneDir = NodeText(xNode.selectSingleNode("donedir"))
    Dim sExt: sExt = NodeText(xNode.selectSingleNode("TargetFileExtension"))        ' shared with main program

    If Len(sHost) = 0 Then
        AppendToLog "ERROR: Accounting Hub FTP host not defined"
        Exit Function
    End If
    If Len(sLogin) = 0 Then
        AppendToLog "ERROR: Accounting Hub FTP login not defined"
        Exit Function
    End If
    If Len(sPattern) = 0 Then
        If Len(sExt) = 0 Then
            sPattern = "*.*"
        Else
            sPattern = "*." & sExt
        End If
    End If
    If Len(sInputDir) = 0 Then
        AppendToLog "ERROR: Input directory not defined"
        Exit Function
    End If
    If Len(sDoneDir) = 0 Then
        AppendToLog "ERROR: Directory for transferred files not defined"
        Exit Function
    End If
    
    If Len(sPassword) = 0 Then
        AHremotehost = "scp://" & sLogin & "@" & sHost & "/"
    Else
        AHremotehost = "scp://" & sLogin & ":" & sPassword & "@" & sHost & "/"
    End If
    AHremotedir = sRemoteDir
    AHpat = sPattern
    AHsslkey = sHostKey
    AHlocaldir = sInputDir
    AHdonedir = sDoneDir
    
    GetConfig = True
End Function

Function GetLogFile(sConfig, sLogFile)
'need a working default!
    GetLogFile = base & sLogFile
    
    Dim sFile
    If InStr(sConfig, "\") = 0 Then
        sFile = base & sConfig
    Else
        sFile = sConfig
    End If
    Dim xmlConfig: Set xmlConfig = CreateObject("MSXML2.DOMDocument")
' load config file
    xmlConfig.async = False
    If Not xmlConfig.load(sFile) Then
        AppendToLog "ERROR: unable to load log4net config file " & sFile & ": error=" & xmlConfig.parseError.reason
        Exit Function
    End If

' get our config section
    Dim xNode
    Set xNode = xmlConfig.selectSingleNode(log4net_filesetting)
    If xNode Is Nothing Then
        AppendToLog "ERROR: can't find log file in " & sFile
        Exit Function
    End If

' read our config parameters
    Dim sLog: sLog = xNode.getAttribute("value")
    If Len(sLog) = 0 Then Exit Function

    sLog = Replace(sLog, "\\", "\")
    Dim iSlash
    iSlash = InStrRev(sLog, "\")
    If iSlash = 0 Then Exit Function
    If Not fso.FolderExists(Left(sLog, iSlash-1)) Then
    msgbox "log dir " & Left(sLog, iSlash-1) & " not found"
        Exit Function
    End If

    sLog = Left(sLog, iSlash) & sLogFile
    GetLogFile = sLog
End Function

' Function to retrieve the test of an XML node, returning an empty string if the node is not found
Function NodeText(x)
    NodeText = "" 
    If x Is Nothing Then Exit Function
    NodeText = x.text
End Function
