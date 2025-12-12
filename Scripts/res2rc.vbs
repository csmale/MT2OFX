' script to convert english strings into stringtable for MT2OFX resource file
Option Explicit
Const ForReading = 1
Const ForWriting = 2

Dim sRcExe: sRcExe = "C:\Program Files\Microsoft Visual Studio\Common\MSDev98\Bin\rc.exe"

Dim sStableHeader: sStableHeader = "STRINGTABLE" & vbCrLf & "LANGUAGE LANG_ENGLISH, SUBLANG_ENGLISH_UK" & vbCrLf & "{"
Dim sStableTrailer: sStableTrailer = "}"
Dim WshShell: Set WshShell = WScript.CreateObject("WScript.Shell")

SetLocale 65001

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim fIn, fOut, sLine, iEq, sNum, sStr

Dim sFileIn, sFileOut

'If WScript.Arguments.Count <> 2 Then
'    MsgBox "Usage: res2rc.vbs <input> <output>"
'    WScript.Quit 2
'End If
'sFileIn = WScript.Arguments(0)
sFileIn = "res\english.lng"
'sFileOut = WScript.Arguments(1)
sFileOut = "bla.rc"
WshShell.CurrentDirectory = "C:\Documents and Settings\Administrator\My Documents\My Projects\MT2OFX"
Set fIn = fso.OpenTextFile(sFileIn, ForReading, False)
Set fOut = fso.OpenTextFile(sFileOut, ForWriting, True)

fOut.WriteLine sStableHeader

If Not fIn.AtEndOfStream Then
    sLine = fIn.ReadLine
    If sLine <> "[MT2OFX Language File]" Then
        MsgBox "Input file is not an MT2OFX Language File" & sLine
        WScript.Quit 3
    End If
End If

Do While Not fIn.AtEndOfStream
    sLine = fIn.ReadLine
    iEq = InStr(sLine, "=")
    If iEq > 0 Then
        sNum = Left(sLine, iEq-1)
        sStr = Mid(sLine, iEq+1)
        sLine.WriteLine sNum & ", """ & Replace(sLine, """", "\""") & """"
    End If
Loop

fOut.WriteLine sStableTrailer
