VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : Logger
'    Project    : MT2OFX
'
'    Description: Error Logging
'
'    Modified   : $Author: Colin $ $Date: 15/06/09 19:25 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/Logger.cls 9     15/06/09 19:25 Colin $"
' $History: Logger.cls $
' 
' *****************  Version 9  *****************
' User: Colin        Date: 15/06/09   Time: 19:25
' Updated in $/MT2OFX
' For transfer to new laptop
'
' *****************  Version 7  *****************
' User: Colin        Date: 7/12/06    Time: 15:07
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 4  *****************
' User: Colin        Date: 6/03/05    Time: 23:42
' Updated in $/MT2OFX
'</CSCC>

Private oFS As Scripting.FileSystemObject
Private oTS As Scripting.TextStream
'This object takes care of logging errors. It can log to the app.log
'or to a local file. The logError function is intended to be used within
'the prompt string in the MsgBox function, but, obviously not from this object.

Private mblnAppLog As Boolean       'Log to application log
Private mblnFileLog As Boolean      'Log to local text file
Private mstrFileName As String      'Text file name
Private mstrPathName As String      'Text file path
Private mblnInit As Boolean         'Text file logging initialized

'windows version checking...
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (VersionInfo As OSVERSIONINFOEX) As Long
    
Const VER_PLATFORM_WIN32s = 0           'Windows 16 bit (obsolesent)
Const VER_PLATFORM_WIN32_WINDOWS = 1    'Windows 95/98/ME
Const VER_PLATFORM_WIN32_NT = 2         'Windows NT/2000/XP
Const OSVERSIONINFOSIZE = 148           'Size of the Older, smaller UDT

Private Type OSVERSIONINFOEX 'Newer UDT with the latest members
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 'everything after this is for the newer version
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    bProductType As Byte
    bReserved As Byte
End Type
Private mlngWinVersion As Long



'LogError returns a concatenated string for use in the MsgBox function prompt
'Kills two lines of code with one bird.....
Public Function LogError(lngNum As Long, strDsc As String, strSrc As String, Optional strProcName As String = "") As String
    Dim strErr As String
    On Error GoTo EH:
    strErr = CStr(lngNum) & " : " & strDsc & "; Source: " & strSrc
    
    LogError = strErr
    
    If mblnFileLog = True Then
        oTS.WriteLine strProcName & "; " & strErr & " Date: " & Format(Now(), "mm/dd/yyyy HH:NN:SS")
    End If
    
    '///allow logging to the App Log on NT/2000/XP machines only///
    If mblnAppLog = True And mlngWinVersion >= VER_PLATFORM_WIN32_NT Then
        App.LogEvent strProcName & "; " & strErr
    
    End If
    
    Exit Function
EH:
    MyMsgBox CStr(Err.Number) & " : " & Err.Description, vbCritical + vbOKOnly, "Logger Error"
End Function
Public Sub Init(blnAppLog As Boolean, blnFileLog As Boolean, Optional strFileName As String = "", Optional strPath As String = "")
    
    On Error GoTo EH:
    
    mblnAppLog = blnAppLog
    mblnFileLog = blnFileLog
    
    If strFileName <> "" Then
        mstrFileName = strFileName
    End If
    If strPath <> "" Then
        If Right$(strPath, 1) <> "\" Then
            mstrPathName = strPath & "\"
        Else
            mstrPathName = strPath
        End If
    End If
    
    'if the specified path does not exist then we will try to create it....
    If oFS.FolderExists(Left$(mstrPathName, Len(mstrPathName) - 1)) = False Then
        oFS.CreateFolder (Left$(mstrPathName, Len(mstrPathName) - 1))
    End If
    
    'create a stream object and create the file if not found...
    Set oTS = oFS.OpenTextFile(mstrPathName & mstrFileName, ForAppending, True)
    
    mblnInit = True
    
        
    Exit Sub
EH:
    MyMsgBox CStr(Err.Number) & " : " & Err.Description, vbCritical + vbOKOnly, "Logger Initialize Error"
End Sub
Public Sub Quit()
    On Error Resume Next
    If mblnInit = True Then
        oTS.Close
        Set oTS = Nothing
    End If
End Sub

Public Property Let UseAppLog(blnLog As Boolean)
    mblnAppLog = blnLog
End Property
Public Property Get UseAppLog() As Boolean
    UseAppLog = mblnAppLog
End Property
Public Property Let UseFileLog(blnLog As Boolean)
    mblnAppLog = blnLog
End Property
Public Property Get UseFileLog() As Boolean
    UseAppLog = mblnAppLog
End Property

Public Property Get LogFile() As String
    LogFile = mstrFileName
End Property
Public Property Get LogFilePath() As String
    LogFilePath = mstrPathName
End Property

Private Sub Class_Initialize()
    On Error GoTo EH
    Set oFS = New Scripting.FileSystemObject
    'set default log file and path
    mstrFileName = "ParcoErrLog.txt"
    mstrPathName = App.Path & "\"
    
    'now determine the platform to see if we can use app.logevent
    Dim VersionInfo As OSVERSIONINFOEX
    VersionInfo.dwOSVersionInfoSize = Len(VersionInfo)
    If GetVersionEx(VersionInfo) = 0 Then 'we chose the wrong UDT
        'take our existing UDT and tell the API that it is shorter than actual
        VersionInfo.dwOSVersionInfoSize = OSVERSIONINFOSIZE
        GetVersionEx VersionInfo
    End If
    mlngWinVersion = VersionInfo.dwPlatformId
    
    
    Exit Sub
EH:
    MyMsgBox CStr(Err.Number) & " : " & Err.Description, vbCritical + vbOKOnly, "Logger Class Initialize Error"
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    Set oFS = Nothing
End Sub


