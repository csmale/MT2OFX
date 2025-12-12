' merge new scriptcat.xml into existing
Option Explicit

Dim isUpgrade
Dim xOld
Dim xNew
Dim sDir: sDir = "C:\Documents and Settings\Administrator\My Documents\My Projects\MT2OFX\"
Const FILE_NOT_FOUND = &H800C0006
Dim ScriptMD5Dict: Set ScriptMD5Dict = CreateObject("Scripting.Dictionary")

' sFile is new scriptcat
Function InitMerge(sFile)
    InitMerge = False
    Set xOld = CreateObject("MSXML.DOMDocument")
    xOld.async = False
    xOld.load sDir & "scriptcat.xml"
    If xOld.parseError.errorCode = 0 Then
        isUpgrade = True
    ElseIf xOld.parseError.errorCode = FILE_NOT_FOUND Then
        isUpgrade = False
        Set xOld = Nothing
    Else
        Exit Function
    End If
    
    Set xNew = CreateObject("MSXML.DOMDocument")
    xNew.async = False
    xNew.load sDir & sFile
    If xNew.parseError.errorCode = 0 Then
        InitMerge = True
    ElseIf xNew.parseError.errorCode = FILE_NOT_FOUND Then
        MsgBox "Unable to load " & sFile & ": File not found"
    Else
        MsgBox "Unable to load " & sFile & ": " & xNew.parseError.reason
    End If
End Function

Function InitMake(sFile)
    Dim sh: Set sh = CreateObject("WScript.Shell")
    Dim sCmd, sLine, iSp, sMD5, sScript
    Dim oExec
    InitMake = False
    Set xNew = CreateObject("MSXML.DOMDocument")
    xNew.async = False
    xNew.load sDir & sFile
    If xNew.parseError.errorCode = 0 Then
        InitMake = True
    ElseIf xNew.parseError.errorCode = FILE_NOT_FOUND Then
        MsgBox "Unable to load " & sFile & ": File not found"
        Exit Function
    Else
        MsgBox "Unable to load " & sFile & ": " & xNew.parseError.reason
        Exit Function
    End If
    sh.CurrentDirectory = sDir
    sCmd = "%comspec% /c md5.exe -l *.vbs"
    Set oExec = sh.Exec(sCmd)
'    Do While oExec.Status = 0
'        WScript.Sleep 100
'    Loop
    If oExec.ExitCode = 0 Then
        Do While Not oExec.StdOut.AtEndOfStream
            sLine = oExec.StdOut.ReadLine
            iSp = InStr(sLine, " ")
            If iSp = 33 Then
                sMD5 = Left(sLine, 32)
        ' two spaces between hash value and script name
                sScript = Mid(sLine, 35)
                ScriptMD5Dict(sScript) = sMD5
            Else
                MsgBox sLine & " space at " & iSp
            End If
        Loop
    Else
        MsgBox "Error from md5.exe: exit code=" & oExec.ExitCode & ": " & oExec.StdOut.ReadAll
    End If
    
End Function

If Not InitMake("_scriptcat.xml") Then
    MsgBox "Unable to set up hashes"
    WScript.Quit 1
End If

DoMake()

WScript.Quit 0

If Not InitMerge("_scriptcat.xml") Then
    MsgBox "Merging script catalog failed."
    WScript.Quit 1
End If

Sub DoMake()
    Dim xScript, xScriptList, xMD5, xScriptScript
    Dim sScript, sMD5
    Set xScriptList = xNew.selectNodes("//mt2ofx/bankscript")
    For Each xScript in xScriptList
        Set xScriptScript = xScript.selectSingleNode("script")
        sScript = xScriptScript.text
        sMD5 = GetMD5(sScript)
        If Len(sMD5) > 0 Then
            Set xMD5 = xScript.getAttributeNode("md5")
            If xMD5 Is Nothing Then
                Set xMD5 = xNew.createAttribute("md5")
            End If
            xMD5.Value = sMD5
            xScriptScript.setAttributeNode(xMD5)
        End If
    Next
    xNew.save "__scriptcat.xml"
End Sub

Function GetMD5(sFile)
    If ScriptMD5Dict.Exists(sFile) Then
        GetMD5 = ScriptMD5Dict(sFile)
    Else
        GetMD5 = ""
    End If
End Function
