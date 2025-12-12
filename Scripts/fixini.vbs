' script to examine MT2OFX.INI files
Option Explicit

Dim sTitle: sTitle = "MT2OFX INI File Troubleshooter"
Const iMinLength = 10000
Dim oShell: Set oShell=CreateObject("Shell.Application")
Dim sAppData: sAppData = oShell.Namespace(26).Self.Path
Dim sProgFiles: sProgFiles = oShell.Namespace(38).Self.Path
Dim sAppDataIni: sAppDataIni = sAppData & "\MT2OFX\MT2OFX.INI"
Dim sProgFilesIni: sProgFilesIni = sProgFiles & "\MT2OFX\MT2OFX.INI"

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim f1, f2

If fso.FileExists(sAppDataIni) Then
	Set f1 = fso.GetFile(sAppDataIni)
Else
	Set f1 = Nothing
End If
If fso.FileExists(sProgFilesIni) Then
	Set f2 = fso.GetFile(sProgFilesIni)
Else
	Set f2 = Nothing
End If

Dim fActive, fOther

If f1 Is Nothing Then
	If f2 Is Nothing Then
		MsgBox "No MT2OFX ini-files found. Is it installed?",,sTitle
	Else
		Set fOther = f1
		Set fActive = f2
	End If
Else
	Set fActive = f1
	Set fOther = f2
End If

MsgBox "Active ini file at: " & fActive.Path & ", " & CStr(fActive.Size) & " bytes, " & FormatDateTime(fActive.DateLastModified, 0),,sTitle

If fActive.Size < iMinLength Then
	If fOther Is Nothing Then
		MsgBox "Active ini file seems too small."
	Else
		If fOther.Size < iMinLength Then
			MsgBox "Active ini file seems too small. There is also a file at " & fOther.Path & " which is also too small.",,sTitle
		Else
			MsgBox "Active ini file seems too small. There is also a file at " & fOther.Path & " which does not seem too small.",,sTitle
		End If
	End If
Else
	MsgBox "Active ini file is OK.",,sTitle
End If
