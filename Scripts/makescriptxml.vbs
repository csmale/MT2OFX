'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 3.1
'
' NAME: 
'
' AUTHOR: Colin Smale , IBM
' DATE  : 31/05/2005
'
' COMMENT: 
'
'==========================================================================

Const outFile = "newscript.xml"

Const baseDir = "C:\Documents and Settings\Administrator\My Documents\My Projects\MT2OFX"

Dim fso
Set fso = CreateObject("scripting.filesystemobject")
Dim fdir
Set fdir = fso.GetFolder(baseDir)
Dim re
Set re = New RegExp
re.Pattern = ".*\.vbs"
re.IgnoreCase = True

Set fout = fso.CreateTextFile(outFile, True)
fout.WriteLine "<?xml>"
fout.WriteLine "<mt2ofx>"
Dim f
For Each f In fdir.Files
	If re.Test(f.Name) Then
		DoFile f
	End If
Next
fout.WriteLine "</mt2ofx>"

Sub DoFile(f)
	fout.WriteLine "<bankscript name=""" & f.Name & """format=""CSV"" id="""">"
	fout.WriteLine "  <script>" & f.Name & "</script>"
	fout.WriteLine "</bankscript>"
End Sub
