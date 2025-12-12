' uninstall and clean up
Option Explicit

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim wsh: Set wsh = CreateObject("WScript.Shell")
Const UninstKey = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MT2OFX_is1"

Dim sProgName: sProgName = ""
Dim sUninstProg: sUninstProg = ""
Dim sQuietUninstProg: sQuietUninstProg = ""
Dim sInstLoc: sInstLoc = ""
Dim sAppDataDir: sAppDataDir = ""
Dim sTmp

On Error Resume Next
sProgName = wsh.RegRead(UninstKey & "\DisplayName")
sUninstProg = wsh.RegRead(UninstKey & "\UninstallString")
sQuietUninstProg = wsh.RegRead(UninstKey & "\QuietUninstallString")
sInstLoc = wsh.RegRead(UninstKey & "\InstallLocation")
On Error Goto 0
sAppDataDir = GetAppDataDir()

If Len(sProgName) = 0 Then
    MsgBox "MT2OFX not installed."
' guess the installation directory...
    sInstLoc = GetProgramFilesDir() & "\MT2OFX"
End If

' normal, but quiet, uninstall
Dim oExec
If Len(sQuietUninstProg) > 0 Then
    Set oExec = wsh.Exec(sQuietUninstProg)
    Do While oExec.Status = 0
        WScript.Sleep 100
    Loop
    If oExec.ExitCode <> 0 Then
        MsgBox "Error during unstallation: " & oExec.ExitCode
    End If
End If

' clear out the registry?

' clean up program directory
If Len(sInstLoc) > 0 Then
    If fso.FolderExists(sInstLoc) Then
msgbox "Installation Folder " & sInstLoc & " exists, deleting..."
        fso.DeleteFolder sInstLoc, True
    End If
End If

' clean up app data directory
If Len(sAppDataDir) > 0 Then
    sTmp = sAppDataDir & "\MT2OFX"
    If fso.FolderExists(sTmp) Then
msgbox "AppData Folder " & sTmp & " exists, deleting..."
        fso.DeleteFolder sTmp, True
    End If
End If

MsgBox "MT2OFX cleanup complete"

Function GetAppDataDir()
    Dim APPLICATION_DATA: APPLICATION_DATA= &H1a&
    Dim objShell, objFolder, objFolderItem
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(APPLICATION_DATA)
    Set objFolderItem = objFolder.Self
    GetAppDataDir = objFolderItem.Path
End Function

Function GetProgramFilesDir()
    Dim PROGRAM_FILES: PROGRAM_FILES = &H26&
    Dim objShell, objFolder, objProgramFiles
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(PROGRAM_FILES)
    Set objProgramFiles = objFolder.Self
    GetProgramFilesDir = objProgramFiles.Path
End Function
' program files is referenced by:  objProgramFiles.Path
