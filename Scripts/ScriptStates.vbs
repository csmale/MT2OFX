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

Const outFile = "crap.htm"

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
fout.WriteLine "<html><head></head><body><table>"
fout.WriteLine "<tr><th>File</th><th>Version</th><th>Modified</th><th>Comment</th></tr>"
Dim f
For Each f In fdir.Files
	If re.Test(f.Name) Then
		DoFile f
	End If
Next
fout.WriteLine "</table></body></html>"

Sub DoFile(f)
	Dim ts
	Set ts = f.OpenAsTextStream(1)
	Dim l
	Dim dVSS
	Dim hFile, hVersion, hDate, hTime, hUser
	Dim m
	Dim re
	Set re = New RegExp
' "$Header: /MT2OFX/La Poste-TSV.vbs 3     14/02/05 22:31 Colin $"
	re.Pattern = ".* ScriptVersion = ""\$Header: /MT2OFX/(.*\.vbs) (\d+) +(\d+/\d+/\d+) (\d+:\d+) (.*)\$"""
	re.IgnoreCase = False
	Do While ts.AtEndOfStream <> True
		l = ts.ReadLine()
		Set m = re.Execute(l)
		If m.Count > 0 Then
			hFile = m(0).SubMatches(0)
			hVersion = m(0).SubMatches(1)
			hDate = m(0).SubMatches(2)
			hTime = m(0).SubMatches(3)
			hUser = m(0).SubMatches(4)
			Exit Do
		End If
		l = ""
	Loop
	ts.Close
	fout.WriteLine "<tr>"
	If f.Name = hFile Then
		fout.WriteLine "<td>" & hFile & "</td>"
	Else
		fout.WriteLine "<td>" & hFile & " (file name=" & f.Name & ")</td>"
	End If
	fout.WriteLine "<td>Version " & hVersion & ", " & hDate & " " & hTime & "</td>"
	fout.WriteLine "<td>" & f.DateLastModified & "</td>"
	If IsDate(hDate & " " & hTime) Then
		dVSS = CDate(hDate & " " & hTime)
		If dVSS < f.DateLastModified Then
			fout.WriteLine "<td><font color='red'><b>Modified!</b></font></td>"
		Else
			fout.WriteLine "<td>&nbsp;</td>"
		End If
	Else
		fout.WriteLine "<td>&nbsp;</td>"
	End If
	fout.WriteLine "</tr>"
End Sub
