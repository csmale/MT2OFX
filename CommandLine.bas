Attribute VB_Name = "CommandLine"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : CommandLine
'    Project    : MT2OFX
'
'    Description: Command line manipulation
'
'    Modified   : $Author: Colin $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/CommandLine.bas 4     30/08/09 13:18 Colin $"
' $History: CommandLine.bas $
' 
' *****************  Version 4  *****************
' User: Colin        Date: 30/08/09   Time: 13:18
' Updated in $/MT2OFX
' fixed small parsing problem in command line
'
' *****************  Version 2  *****************
' User: Colin        Date: 7/12/06    Time: 14:53
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2

'</CSCC>

Private Type MungeLong
    X As Long
    Dummy As Integer
End Type

Private Type MungeInt
    XLo As Integer
    XHi As Integer
    Dummy As Integer
End Type

Private Declare Function GetWinCommandLine Lib "kernel32" Alias "GetCommandLineW" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, _
    ByVal Length As Long)
Private Declare Function StrLenW Lib "kernel32.dll" Alias "lstrlenW" (ByVal Ptr As Long) As Long
Private Declare Function CommandLineToArgv Lib "shell32" Alias "CommandLineToArgvW" (ByVal lpCmdLine As Long, _
    ByRef pNumArgs As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (dest As Any, src As Any, ByVal size&)
Private Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" (RetVal As Byte, ByVal Ptr As Long) As Long
Private Declare Function PtrToInt Lib "kernel32" Alias "lstrcpynW" (RetVal As Any, ByVal Ptr As Long, ByVal nCharCount As Long) As Long
Private Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal Ptr As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long


'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetCommandLine
' Description:       Retrieve Command Line without argv(0)
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       8/21/2006-08:23:03
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetCommandLine() As String

    Dim lpStr As Long, i As Long
    Dim buffer As String
        
    ' get a pointer to the command line
    lpStr = GetWinCommandLine()
    ' copy into a local buffer
    i = StrLenW(lpStr)
    buffer = String(i, vbNullChar)
    CopyMemory ByVal StrPtr(buffer), ByVal lpStr, i * 2
    DBCSLog buffer, "From GetCommandLineW"
    GetCommandLine = buffer
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ParseCommandLine
' Description:       Splits command line into Argv/Argc array of parameters
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       8/23/2006-08:37:13
'
' Parameters :       sCommandLine (String) As String()
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ParseCommandLine(sCommandLine As String) As String()
    Dim sCommandLineW As String
    Dim BufPtr As Long
    Dim lNumArgs As Long
    Dim i As Long
    Dim lRes As Long
    Dim TempPtr As MungeLong
    Dim TempStr As MungeInt
    Dim ArgArray(512) As Byte
    Dim Arg As String
    Dim Args() As String
    Dim iNull As Long
    Dim iPtr As Long
    Dim iLen As Long
    Dim iArgPtr As Long
    
'    sCommandLineW = StrConv(sCommandLine, vbUnicode)
    sCommandLineW = sCommandLine
    DBCSLog sCommandLineW, "Unicode command line before parse"
    BufPtr = CommandLineToArgv(StrPtr(sCommandLineW), lNumArgs)
    If lNumArgs = 0 Then
        ParseCommandLine = ParseW98(sCommandLine)
        Exit Function
    End If
    If BufPtr = 0 Then
        ShowError "ParseCommandLine", "Unable to retrieve command line"
        ParseCommandLine = Array()
        Exit Function
    End If
    ReDim Args(lNumArgs - 1)

' BufPtr points to an array of pointers to null-terminated unicode strings
    iPtr = BufPtr
    For i = 1 To lNumArgs
        CopyMemory ByVal VarPtr(lRes), ByVal iPtr, 4&
        iLen = StrLen(lRes)
        Arg = Space$(iLen)
        CopyMemory StrPtr(Arg), lRes, iLen * 2
        iPtr = iPtr + 4
        Args(i - 1) = Arg
        DBCSLog Arg, "Argv(" & i - 1 & ")"
    Next

    Call GlobalFree(BufPtr)
    ParseCommandLine = Args
End Function

Private Function ParseW98(sCommandLine As String) As String()
    Dim Args() As String
    Dim Arg As String
    Dim iTerm As Long
    Dim sLine As String
    Dim argc As Long
    
    ReDim Args(0)
    argc = 0
    sLine = Trim$(sCommandLine)
    
    Do While Len(sLine) > 0
        If Left$(sLine, 1) = """" Then
            iTerm = InStr(2, sLine, """")
            If iTerm > 0 Then
                Arg = Mid$(sLine, 2, iTerm - 2)
                sLine = Trim$(Mid$(sLine, iTerm + 1))
            Else
                Arg = Trim$(Mid$(sLine, 2))
                sLine = ""
            End If
        Else
            iTerm = InStr(2, sLine, " ")
            If iTerm > 0 Then
                Arg = Left$(sLine, iTerm - 1)
                sLine = Trim$(Mid$(sLine, iTerm + 1))
            Else
                Arg = Trim$(sLine)
                sLine = ""
            End If
        End If
        ReDim Preserve Args(argc)
        Args(argc) = Arg
        DBCSLog Arg, "[W98] Argv(" & argc & ")"
        argc = argc + 1
    Loop
    ParseW98 = Args
End Function
