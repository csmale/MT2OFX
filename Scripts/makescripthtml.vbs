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

Const cssStyleSheet = "_themes/mdshapes/mdsh1011.css"
Const baseDir = "."
Const VSSBase = "\\monster\vss60\srcsafe.ini"
Const VSSProject = "$/MT2OFX"
Const VSSUser = "colin"
Const VSSPass = "speak1x"

Dim vssDB
Set vssDB=CreateObject("SourceSafe")

Const sScriptCat = "scriptcat.xml"
Dim xmlCat
Set xmlCat = CreateObject("MSXML.DOMDocument")
xmlCat.Async = False
If Not xmlCat.Load(baseDir & "\" & sScriptCat) Then
	MsgBox "Unable to load script catalog: " & xmlCat.ParseError.Reason & vbCrLf _
	& "at position " & xmlCat.ParseError.linepos & " in " & xmlCat.ParseError.srcText
End If

Const URLBase = ""
Const FlagURLBase = "../images/flags/"

vssDB.Open VSSBase, VSSUser, VSSPass
Dim vssRoot
Set vssRoot = vssDB.VSSItem(VSSProject)
Dim vssItem
Dim vssVer
Dim re
Dim fso
Set fso = CreateObject("scripting.filesystemobject")
Dim fOut
Dim xItem

Set fOut = fso.CreateTextFile(baseDir & "\scripts.htm", True)

Set re = New RegExp
re.Pattern = ".*\.vbs"
re.IgnoreCase = True

DoFileHeader
For Each vssItem In vssRoot.Items
	If re.Test(vssItem.Name) Then
		DoFile vssItem
	End If
'	Exit For
Next
DoFileTrailer
fOut.Close

Sub DoFileHeader()
	fOut.Write "<html><head>"
    If Len(cssStyleSheet) > 0 Then
        fOut.Write "<link rel='stylesheet' type='text/css' href='" & cssStyleSheet & "'>"
    End If
    fOut.WriteLine "</head><body>"
	fOut.WriteLine "<table>" _
		& "<th align='left'>Bank</th>" _
		& "<th align='left'>Format</th>" _
		& "<th align='center'>Country</th>" _
		& "<th align='left'>Version</th>" _
		& "<th align='left'>Date</th>" _
      & "<th align='left'>Script</th>"
'		& "<th align='left'>File Date</th>" _
'		& "<th align='left'>Notes</th>"
End Sub

Function GetText(xNode)
	If xNode Is Nothing Then
		GetText = ""
	Else
		GetText = xNode.text
	End If
End Function

Sub DoFile(vssItem)
	Dim xmlItem
	Set xmlItem = xmlCat.SelectSingleNode("/mt2ofx/bankscript[script/text()='" & vssItem.Name & "']")
	Dim xmlNode
	Dim sRegion: sRegion = ""
	Dim sBank: sBank = ""
	Dim sFormat: sFormat = ""
	Dim f, dFile, sFileDate
	Dim sNotes: sNotes = "&nbsp;"
   Dim sCountry: sCountry = ""

	If xmlItem Is Nothing Then
        Exit Sub
   Else
		sRegion = GetText(xmlItem.SelectSingleNode("region"))
		sBank = GetText(xmlItem.SelectSingleNode("@name"))
		sFormat = GetText(xmlItem.SelectSingleNode("@format"))
      sCountry = GetText(xmlCat.selectSingleNode("/mt2ofx/region[@id='" & sRegion & "']/@en"))
	End If
	If fso.FileExists(vssItem.Name) Then
		Set f = fso.GetFile(vssItem.Name)
		dFile = f.DateLastModified
		sFileDate = FormatDateTime(dFile, vbShortDate)
	Else
		Set f = Nothing
		sFileDate = "&nbsp;"
		sNotes = "File not found."
	End If
	
	Dim vssVersions
	Set vssVersions = vssItem.Versions(0)
	For Each vssVer In vssVersions
		If Left(vssVer.Action, 5) <> "Label" Then
'	MsgBox TypeName(vssVer)
'	fOut.WriteLine vssItem.Name & " at version " & CStr(vssVer.VersionNumber) & " on " & CStr(vssVer.Date)
		If Len(sBank) = 0 Then sBank = "&nbsp;"
		If Len(sFormat) = 0 Then sFormat = "&nbsp;"
		If Len(sRegion) = 0 Then
			sRegion = "&nbsp;"
		Else
			sRegion = "<img src='" & FlagURLBase & sRegion & ".gif' alt='" & sRegion & "' height=12 width=18 />"
		End If
		If Not (f Is Nothing) Then
			If vssVer.Date = dFile Then
				sNotes = "Pristine"
			ElseIf vssVer.Date < dFile Then
				sNotes = "Modified"
			End If
		End If
		fOut.WriteLine "<tr>" _
			& "<td>" & HTMLEncode(sBank) & "</td>" _
			& "<td>" & HTMLEncode(sFormat) & "</td>" _
			& "<td align='center'>" & sRegion & HTMLEncode(sCountry) & "</td>" _
			& "<td align='center'>" & CStr(vssVer.VersionNumber) & "</td>" _
			& "<td>" & FormatDateTime(vssVer.Date, vbShortDate) & "</td>" _
			& "<td>" & HTMLEncode(vssItem.Name) & "</td></tr>"
'			& "<td><a href='" & HTMLEncode(URLBase & vssItem.Name) & "'>" & HTMLEncode(vssItem.Name) & "</a></td></tr>"
'			& "<td>" & sFileDate & "</td>" _
'			& "<td>" & HTMLEncode(sNotes) & "</td>" _
'			& "</tr>"
		Exit For
		End If
	Next
End Sub

Sub DoFileTrailer()
	fOut.WriteLine "</table></body></html>"
End Sub

Function HTMLEncode(sIn)
    Dim c, i, sTmp, ic, sChar
    For i=1 To Len(sIn)
        c = Mid(sIn, i, 1)
        ic = AscW(c)
        If ic<32 Or ic>126 Then
            sChar = "&#x" & Hex(ic) & ";"
        ElseIf c="&" Then
            sChar = "&amp;"
        ElseIf c="<" Then
            sChar = "&lt;"
        ElseIf c=">" then
            sChar = "&gt;"
        Else
            sChar = c
        End If
        sTmp = sTmp & sChar
    Next
    HTMLEncode = sTmp
End Function
