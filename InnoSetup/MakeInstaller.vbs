'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 2007
'
' NAME: 
'
' AUTHOR: Colin Smale
' DATE  : 28/08/2008
'
' COMMENT: 
'
'==========================================================================

Option Explicit

Const csCatalog = "..\scriptcat.xml"
Const csInput = "mt2ofx.iss"
Const csOutput = "mt2ofx_new.iss"
Const csXpathScripts = "/mt2ofx/bankscript/script"
Const csStartScriptArea = "; <<START SCRIPTS>>"
Const csEndScriptArea = "; <<END SCRIPTS>>"
Const csSourceDir = "SourceDir="
Const ForReading = 1
Const ForWriting = 2
Const TristateFalse = 0

On Error Resume Next

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim xIn: Set xIn = CreateObject("Msxml2.DOMDocument")
If Not xIn.load(csCatalog) Then
	MsgBox "Unable to load " & csCatalog
	WScript.Quit 1
End If

Dim xScripts: Set xScripts = xIn.selectNodes(csXpathScripts)
Dim aScripts: Set aScripts = CreateObject("Scripting.Dictionary")
Dim xOne
Dim sScript
If xScripts Is Nothing Then
	MsgBox "Catalog contains no scripts!"
	WScript.Quit	
End If

For Each xOne In xScripts
	sScript = xOne.text
'	MsgBox "script: " & sScript
	aScripts(sScript) = "1"
Next

' now we have a complete list of the scripts required in aScripts
Dim fIn: Set fIn = fso.OpenTextFile(csInput,ForReading,False)
If fIn Is Nothing Then
	MsgBox "Unable to open input installer script"
	WScript.Quit 2
End If
Dim fOut: Set fOut = fso.OpenTextFile(csOutput, ForWriting, True, TristateFalse)
If fOut Is Nothing Then
	MsgBox "Unable to open output installer script: " & Err.Description
	WScript.Quit 3
End If
Dim wsh: Set wsh = WScript.CreateObject("WScript.Shell")
Dim sSourceDir: sSourceDir = fso.GetParentFolderName(wsh.CurrentDirectory)
Dim sLine
Do While fIn.AtEndOfStream <> True
	sLine = fIn.ReadLine
	If Left(sLine, 10) = csSourceDir Then
		sLine = csSourceDir & sSourceDir
	End If
	fOut.WriteLine sLine
    If sLine = csStartScriptArea Then
    	Do While fIn.AtEndOfStream <> True
    		sLine = fIn.ReadLine
    		If sLine = csEndScriptArea Then
    			Exit Do
    		End If
    	Loop
      	Exit Do
    End If
Loop
Dim k
' Source: ABN Amro HomeNet.vbs; DestDir: {app}; Flags: ignoreversion
For Each k In aScripts.Keys
	fOut.WriteLine "Source: " & k & "; DestDir: {app}; Flags: ignoreversion"
Next
fOut.WriteLine csEndScriptArea
Do While fIn.AtEndOfStream <> True
	fOut.WriteLine fIn.ReadLine
Loop

WScript.Quit 0
