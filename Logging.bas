Attribute VB_Name = "Logging"
Option Explicit

' $Header: /MT2OFX/Logging.bas 16    30/01/11 11:15 Colin $

Private oLogFile As TextStream
Const bLogUnicode As Boolean = False

' for DBCS logging
'Public iDBCSLogFile As Integer
Const sDBCSLogFile As String = "dbcslog.txt"
Const sDebugLogFile As String = "debuglog.txt"
Private oDBCSLogFile As TextStream
Private bDBCSLogUnicode As Boolean

' for event log usage
Private Declare Function RegisterEventSource Lib "advapi32.dll" Alias _
        "RegisterEventSourceA" (ByVal lpUNCServerName As String, _
        ByVal lpSourceName As String) As Long
Private Declare Function DeregisterEventSource Lib "advapi32.dll" ( _
        ByVal hEventLog As Long) As Long
Private Declare Function ReportEvent Lib "advapi32.dll" Alias _
      "ReportEventA" ( _
        ByVal hEventLog As Long, ByVal wType As Integer, _
        ByVal wCategory As Integer, ByVal dwEventID As Long, _
        ByVal lpUserSid As Any, ByVal wNumStrings As Integer, _
        ByVal dwDataSize As Long, plpStrings As Long, _
        lpRawData As Any) As Boolean
      
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        hpvDest As Any, hpvSource As Any, _
        ByVal cbCopy As Long)
Private Declare Function GlobalAlloc Lib "kernel32" ( _
         ByVal wFlags As Long, _
         ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" ( _
         ByVal hMem As Long) As Long

Public Const EVENTLOG_SUCCESS = 0
Public Const EVENTLOG_ERROR_TYPE = 1
Public Const EVENTLOG_WARNING_TYPE = 2
Public Const EVENTLOG_INFORMATION_TYPE = 4
Public Const EVENTLOG_AUDIT_SUCCESS = 8
Public Const EVENTLOG_AUDIT_FAILURE = 10


'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       MessageToScreen
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       18/03/2005-21:18:57
'
' Parameters :       Text (String)
'                    Title (String)
'                    Flags (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub MessageToScreen(Text As String, Title As String, Flags As Long)
    MyMsgBox Text, Flags, Title
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       MessageToFile
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       18/03/2005-21:19:06
'
' Parameters :       Text (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub MessageToFile(Text As String)
    Dim sTmp As String
    Dim oFSO As New FileSystemObject
    On Error GoTo bad_open
    If oLogFile Is Nothing Then
        If Cfg.LogFile = "" Then Exit Sub
        If PathIsRelative(Cfg.LogFile) Then
' 20061129 CS: changed to AppData for Vista compatibility
            sTmp = Cfg.AppDataPath & "\" & Cfg.LogFile
        Else
            sTmp = Cfg.LogFile
        End If
        Set oLogFile = oFSO.OpenTextFile(sTmp, ForAppending, True, IIf(bLogUnicode, TristateTrue, TristateFalse))
'        Open sTmp For Append Access Write As iLogFile
    End If
    sTmp = Format(Now(), "General Date")
    sTmp = sTmp & " " & Text
' if we are logging in ANSI/ASCII mode, WriteLine barfs if there are any "impossible" characters
    If Not bLogUnicode Then
        sTmp = StrConv(StrConv(sTmp, vbFromUnicode), vbUnicode)
    End If
    oLogFile.WriteLine sTmp
    Exit Sub
bad_open:
    MessageToScreen "(" & CStr(Err.Number) & ") :" & Err.Description, "Log File Error", vbOKOnly + vbCritical
    If Not oLogFile Is Nothing Then
        oLogFile.Close
        Set oLogFile = Nothing
    End If
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       LogMessage
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       18/03/2005-21:18:05
'
' Parameters :       ToScreen (Boolean)
'                    ToFile (Boolean)
'                    Text (String)
'                    Title (String = "")
'                    Flags (Long = vbOK)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub LogMessage(ToScreen As Boolean, ToFile As Boolean, Text As String, _
    Optional Title As String = "MT2OFX", Optional Flags As Long = vbOKOnly)
    Dim sTmp As String
    If Len(Title) = 0 Then
        sTmp = Text
    Else
        sTmp = Title & ": " & Text
    End If
    Debug.Print sTmp
    If ToFile Then
        MessageToFile sTmp
    End If
    If ToScreen Then
        MessageToScreen Text, Title, Flags
    End If
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CloseLogFile
' Description:       close the log file if it is open
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       09/02/2005-22:20:14
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub CloseLogFile()
    On Error Resume Next
    If Not oLogFile Is Nothing Then
        oLogFile.Close
        Set oLogFile = Nothing
    End If
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ShowError
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       13/10/2004-10:18:54
'
' Parameters :       sTitle (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub ShowError(sModule As String, Optional sAdditional As String = "", Optional sTitle As String = "")
    Dim sMsg As String
    Dim sTit As String
    If sTitle = "" Then
        sTit = "MT2OFX Error"
    Else
        sTit = sTitle
    End If
    sMsg = Err.Description _
        & vbCrLf & "In " & sModule _
        & vbCrLf & "From " & Err.Source _
        & vbCrLf & "(Error number: " & CStr(Err.Number) & ")"
    If Len(sAdditional) > 0 Then
        sMsg = sMsg & vbCrLf & sAdditional
    End If
    LogMessage False, True, sMsg, sTit
    MyMsgBox sMsg, vbOKOnly + vbCritical, sTit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ShowDllError
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       22/10/2004-22:28:57
'
' Parameters :       sModule (String)
'                    sAdditional (String = "")
'                    sTitle (String = "")
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub ShowDllError(sModule As String, Optional sAdditional As String = "", Optional sTitle As String = "")
    Dim sMsg As String
    Dim sTit As String
    If sTitle = "" Then
        sTit = "MT2OFX Error"
    Else
        sTit = sTitle
    End If
    Dim iErrNum As Long
    iErrNum = Err.LastDllError
    sMsg = ReturnAPIError(iErrNum) _
        & vbCrLf & "In " & sModule _
        & vbCrLf & "(Error number: " & CStr(iErrNum) & ")"
    If Len(sAdditional) > 0 Then
        sMsg = sMsg & vbCrLf & sAdditional
    End If
    LogMessage False, True, sMsg, sTit
    MyMsgBox sMsg, vbOKOnly + vbCritical, sTit
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DBCSLog
' Description:       Write log record for DBCS debugging
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       8/17/2006-20:19:52
'
' Parameters :       sPath (String)
'                    sText (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub DBCSLog(sPath As String, sText As String)
    Dim c As String
    Dim i As Integer
    Dim oFSO As New FileSystemObject
    Dim sLine As String
    Dim oVI As OSVERSIONINFOEX
    Dim bUnicode As Tristate
    
    On Error GoTo baleout
    
    If Not Cfg.DebugDBCS Then
        Exit Sub
    End If
    
    If oDBCSLogFile Is Nothing Then
        If sDBCSLogFile <> "" Then
            oVI.dwOSVersionInfoSize = Len(oVI)
            If GetVersionEx(oVI) <> 0 Then
                bDBCSLogUnicode = (oVI.dwMajorVersion >= 5)
            Else
                bDBCSLogUnicode = False
            End If
            bDBCSLogUnicode = False
            If bDBCSLogUnicode Then
                bUnicode = TristateTrue
            Else
                bUnicode = TristateFalse
            End If
            Set oDBCSLogFile = oFSO.OpenTextFile(Cfg.AppDataPath & "\" & sDBCSLogFile, ForWriting, True, bUnicode)
        End If
    End If
    
    If Not oDBCSLogFile Is Nothing Then
        oDBCSLogFile.WriteLine p(Replace(sText, vbNullChar, ".") & ":", bDBCSLogUnicode)
        For i = 1 To Len(sPath)
            c = Mid$(sPath, i, 1)
            If i Mod 16 = 1 Then
                sLine = Right$("0000" & Hex(i - 1), 4) & ": "
            End If
            sLine = sLine & Right$("0000" & Hex(AscW(c)), 4) & " "
            If i Mod 16 = 0 Then
                oDBCSLogFile.WriteLine sLine
                sLine = ""
            End If
        Next
        If Len(sLine) > 0 Then
            oDBCSLogFile.WriteLine sLine
        End If
        oDBCSLogFile.WriteBlankLines 1
        oDBCSLogFile.WriteLine p(Replace(sPath, vbNullChar, "."), bDBCSLogUnicode)
        oDBCSLogFile.WriteBlankLines 1
    End If
    Exit Sub
    
baleout:
    Set oDBCSLogFile = Nothing
End Sub

Private Function p(sLine As String, bUnicode As Boolean) As String
    If bUnicode Then
        p = sLine
    Else
        p = StrConv(StrConv(sLine, vbFromUnicode), vbUnicode)
    End If
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CloseDBCSLog
' Description:       Close DBCS Log File
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       8/21/2006-14:34:50
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub CloseDBCSLog()
    If Not oDBCSLogFile Is Nothing Then
        oDBCSLogFile.Close
        Set oDBCSLogFile = Nothing
    End If
End Sub

Public Sub DebugLog(sText As String, Optional iLevel As logLevels = logNORMAL)
    Dim oFSO As Scripting.FileSystemObject
    Dim oLog As Scripting.TextStream
    Dim sTime As String, sTmp As String
    If Cfg.LogLevel >= iLevel Then
        sTime = Format(Now(), "dd-mm-yyyy hh:MM:ss")
        sTmp = sTime & ": " & sText
        Set oFSO = New Scripting.FileSystemObject
        Set oLog = oFSO.OpenTextFile(Cfg.AppDataPath & "\" & sDebugLogFile, Scripting.ForAppending, True, False)
        oLog.WriteLine sTmp
        oLog.Close
    End If
End Sub

Public Sub RecordEvent(iType As Integer, iEventID As Long, sMessage As String)

    Dim lRetv As Long
    Dim hEventLog As Long
    Dim hMsg As Long
    Dim hDat As Long
    Dim cbStringSize As Long
    Dim cbDataSize As Long
    Dim iNumStrings As Integer
    
    iNumStrings = 1
    
    '   open the event log
    hEventLog = RegisterEventSource("", "MT2OFX")
    
    '   calculate the amount of memory needed for the string
    cbStringSize = LenB(sMessage) + 1
    '   calculate the amount of memory needed for the string as data
    cbDataSize = Len(sMessage)
    
    '   allocate the memory needed and initialize it to 0's
    hMsg = GlobalAlloc(&H40, cbStringSize)
    hDat = GlobalAlloc(&H40, cbDataSize)
    
    '   copy the string to the memory allocated
    '   be sure to append a \0 to the string
    CopyMemory ByVal hMsg, sMessage & vbNullString, cbStringSize
    '   copy the string to the data memory allocated
    CopyMemory ByVal hDat, ByVal sMessage, cbDataSize
    
    '   log the event
    lRetv = ReportEvent(hEventLog, iType, 0, ByVal iEventID, ByVal 0&, 1, cbDataSize, ByVal hMsg, ByVal hDat)
    
    '   debug information
    '   you could use FormatMessage and GetLastError here
    If lRetv = 0 Then Debug.Print "Error"

    '   free up the memory used
    Call GlobalFree(hMsg)
    Call GlobalFree(hDat)
    '   free up the pointer to the event log
    DeregisterEventSource (hEventLog)

End Sub

