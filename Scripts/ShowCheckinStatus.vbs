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

Const outFile = "CheckinStatus.htm"

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

Const sVBProject = "MT2OFX.vbp"

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

Set fOut = fso.CreateTextFile(baseDir & "\" & outFile, True)

Set re = New RegExp
re.Pattern = ".*\.vbs"
re.IgnoreCase = True
Dim re2
Set re2 = New RegExp
re2.Pattern = ".*\.bas|.*\.cls|.*\.frm"
re2.IgnoreCase = True

DoFileHeader
For Each vssItem In vssRoot.Items
    If re.Test(vssItem.Name) Then
        DoScriptFile vssItem
    ElseIf re2.Test(vssItem.Name) Then
        DoVBFile vssItem
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
		& "<th align='left'>VDate</th>" _
		& "<th align='left'>FDate</th>" _
      & "<th align='left'>Script</th>" _
		& "<th align='left'>Notes</th>"
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

Sub DoScriptFile(vssItem)
	Dim xmlItem
	Set xmlItem = xmlCat.SelectSingleNode("/mt2ofx/bankscript[script/text()='" & vssItem.Name & "']")
	Dim xmlNode
	Dim sRegion: sRegion = ""
	Dim sBank: sBank = ""
	Dim sFormat: sFormat = ""
	Dim f, dFile, sFileDate, dVersion
	Dim sNotes: sNotes = ""
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
		sFileDate = ""
		sNotes = "File not found."
	End If
	
	Dim vssVersions
	Set vssVersions = vssItem.Versions(0)
	For Each vssVer In vssVersions
'		If Left(vssVer.Action, 5) <> "Label" Then
'	MsgBox TypeName(vssVer)
'	fOut.WriteLine vssItem.Name & " at version " & CStr(vssVer.VersionNumber) & " on " & CStr(vssVer.Date)
		If Len(sRegion) > 0 Then
			sRegion = "<img src='" & FlagURLBase & sRegion & ".gif' alt='" & sRegion & "' height=12 width=18 />"
		End If
        dVersion = vssVer.Date
        If Not (f Is Nothing) Then
            sNotes = CompareDates(dFile, dVersion)
        Else
            sNotes = ""
        End If
		fOut.WriteLine "<tr>" _
			& "<td>" & HTMLEncode2(sBank) & "</td>" _
			& "<td>" & HTMLEncode2(sFormat) & "</td>" _
			& "<td align='center'>" & sRegion & HTMLEncode2(sCountry) & "</td>" _
			& "<td align='center'>" & CStr(vssVer.VersionNumber) & "</td>" _
			& "<td>" & FormatDateTime(vssVer.Date, vbGeneralDate) & "</td>" _
			& "<td>" & FormatDateTime(dFile, vbGeneralDate) & "</td>" _
			& "<td>" & HTMLEncode2(vssItem.Name) & "</td>" _
			& "<td>" & HTMLEncode2(sNotes) & "</td></tr>"
'			& "<td><a href='" & HTMLEncode(URLBase & vssItem.Name) & "'>" & HTMLEncode(vssItem.Name) & "</a></td></tr>"
'			& "<td>" & sFileDate & "</td>" _
'			& "<td>" & HTMLEncode2(sNotes) & "</td>" _
'			& "</tr>"
		Exit For
'		End If
	Next
End Sub

Sub DoVBFile(vssItem)
	Dim sRegion: sRegion = ""
	Dim sBank: sBank = ""
	Dim sFormat: sFormat = ""
	Dim f, dFile, sFileDate, dVersion
	Dim sNotes: sNotes = ""
   Dim sCountry: sCountry = ""

	If fso.FileExists(vssItem.Name) Then
		Set f = fso.GetFile(vssItem.Name)
		dFile = f.DateLastModified
		sFileDate = FormatDateTime(dFile, vbShortDate)
	Else
		Set f = Nothing
		sFileDate = ""
		sNotes = "File not found."
	End If
	
	Dim vssVersions
	Set vssVersions = vssItem.Versions(0)
	For Each vssVer In vssVersions
'		If Left(vssVer.Action, 5) <> "Label" Then
'	MsgBox TypeName(vssVer)
'	fOut.WriteLine vssItem.Name & " at version " & CStr(vssVer.VersionNumber) & " on " & CStr(vssVer.Date)
		If Len(sRegion) > 0 Then
			sRegion = "<img src='" & FlagURLBase & sRegion & ".gif' alt='" & sRegion & "' height=12 width=18 />"
		End If
        dVersion = vssVer.Date
        If Not (f Is Nothing) Then
            sNotes = CompareDates(dFile, dVersion)
        Else
            sNotes = ""
        End If
		fOut.WriteLine "<tr>" _
			& "<td>" & HTMLEncode2(sBank) & "</td>" _
			& "<td>" & HTMLEncode2(sFormat) & "</td>" _
			& "<td align='center'>" & sRegion & HTMLEncode2(sCountry) & "</td>" _
			& "<td align='center'>" & CStr(vssVer.VersionNumber) & "</td>" _
			& "<td>" & FormatDateTime(vssVer.Date, vbGeneralDate) & "</td>" _
			& "<td>" & FormatDateTime(dFile, vbGeneralDate) & "</td>" _
			& "<td>" & HTMLEncode2(vssItem.Name) & "</td>" _
			& "<td>" & HTMLEncode2(sNotes) & "</td></tr>"
'			& "<td><a href='" & HTMLEncode(URLBase & vssItem.Name) & "'>" & HTMLEncode(vssItem.Name) & "</a></td></tr>"
'			& "<td>" & sFileDate & "</td>" _
'			& "<td>" & HTMLEncode2(sNotes) & "</td>" _
'			& "</tr>"
		Exit For
'		End If
	Next
End Sub

Function CompareDates(ByVal dFile, ByVal dVersion)
    Dim sTmp
    Dim dDiff
    dDiff = Abs(DateDiff("s", dFile, dVersion))
    If dDiff < 5 Then
        sTmp = "Pristine"
    ElseIf dDiff > 3595 And dDiff < 3605 Then
        sTmp = "Pristine"
    Else
        If dFile < dVersion Then
            sTmp = "File is old version"
        Else
            sTmp = "Modified (" & dDiff & ")"
        End If
    End If
    CompareDates = sTmp
End Function

Sub DoFileTrailer()
	fOut.WriteLine "</table></body></html>"
End Sub

Function HTMLEncode2(sIn)
    If Len(sIn) = 0 Then
        HTMLEncode2 = "&nbsp;"
    Else
        HTMLEncode2 = HTMLEncode(sIn)
    End If
End Function

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
