Attribute VB_Name = "LongFileNames"
Option Explicit

' $Header: /MT2OFX/LongFileNames.bas 8     24/07/10 10:45 Colin $

' Thanks to Patrice Goyer (patrice.goyer@goelis.com)

' If we got the file name through "send to" we only got the short (DOS) name
' This may be troublesome, e.g. if we use the filename to generate
' another filename

' We can use the API function FindFirstFile() to get the full name
' Win32 API
Private Const MAX_PATH As Long = 520

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName(1 To MAX_PATH) As Byte
        cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileW" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameW" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameW" (ByVal lpShortName As Long, ByVal lpLongName As Long, ByVal nBufLen As Long) As Long

Public Const GENERIC_READ = &H80000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal NoSecurity As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long
Public Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long

Private Declare Function StrLenW Lib "kernel32.dll" Alias "lstrlenW" (ByVal Ptr As Long) As Long

Public Function ShortFileNameToLong(FileName As String) As String
    Dim i As Long
    Dim sTmpIn As String
    Dim lErr As Long
    Dim sTmp As String
' handle long paths better
    i = GetLongPathName(StrPtr(FileName), 0&, 0&)
    If i = 0 Then
        lErr = GetLastError()
        ShowError "GetLongPathName", "Error " & CStr(lErr) & " from GetLongPathName('" & FileName & "',0,0)", ReturnAPIError(lErr)
        ShortFileNameToLong = FileName
        Exit Function
    End If
    sTmp = String$(i, vbNullChar)
    i = GetLongPathName(StrPtr(FileName), StrPtr(sTmp), Len(sTmp))
    If i = 0 Then
        lErr = GetLastError()
        ShowError "GetLongPathName", "Error " & CStr(lErr) & " from GetLongPathName('" & FileName & "',x,y)", ReturnAPIError(lErr)
        ShortFileNameToLong = FileName
        Exit Function
    End If
    ShortFileNameToLong = Left(sTmp, InStr(sTmp, vbNullChar) - 1)
End Function

Public Function ShortFileNameToLongz(FileName As String) As String
Dim hFind As Long
Dim fileData As WIN32_FIND_DATA
Dim fn As String, fp As String
Dim s As String, sFile, lg As Long
Dim lErr As Long
Dim i As Long

    On Error Resume Next

' First use FindFirstFile() to find the full name
' Pitfall: FindFirstFile() does not return the path name.
' It just returns the base name
' So we first split the file name to recall the file's location

    ExtractFileNameAndPath FileName, fn, fp

' MsgBox "path: " & fp & ", name: " & fn
' DBCS: ok here
    DBCSLog fp, "ShortFileNameToLong: Path"
    DBCSLog fn, "ShortFileNameToLong: Filename"

    hFind = FindFirstFile(StrPtr(FileName), fileData)
    If hFind = INVALID_HANDLE_VALUE Then
        lErr = GetLastError()
        ShowError "ShortFileNameToLong", "Error " & CStr(lErr) & " from FindFirstFile('" & FileName & "')", ReturnAPIError(lErr)
        Exit Function
    End If
    
    DBCSLog FileName, "Back from FindFrstFile"
    
    ' Then we paste back the file name with the path
    i = StrLenW(VarPtr(fileData.cFileName(1)))
    FileName = Left(fileData.cFileName, i)
    If fp = "" Then
'        FileName = stripNullChars(fileData.cFileName)
    Else
        FileName = fp & "\" & FileName
    End If

    DBCSLog FileName, "Recombined"
    
    ' free the "find" handle
    Call FindClose(hFind)

    ' Finally we use the GetFullPathName() API function to get
    ' the long file name with its path

    ' Note: we make a first call with an empty string and
    ' length=0 in order to get the length of the full name

    s = ""
    lg = 0
    sFile = ""
    lg = GetFullPathName(FileName, lg, s, sFile)

    s = Space$(lg) ' get enough memory to contain the file name
    sFile = Space$(lg)

    ' now we make the actual call to get the value
    lg = GetFullPathName(FileName, lg, s, sFile)
' lg is unreliable in DBCS environments but the string is guaranteed to be null-terminated
    DBCSLog s, "GetFullPathName returns " & CStr(lg) & " chars"
    lg = InStr(s, vbNullChar) - 1
    
    ' return with the long path name
    ShortFileNameToLongz = Left$(s, lg)  ' strip off the trailing '\0' of the C string

End Function

' Split the filename to get basename and file location
' Note: this could also be done using common dialog control smartly
Private Sub ExtractFileNameAndPath(fullName As String, fName As String, fPath As String)
Dim i As Integer, l As Integer, c As String
Dim found As Boolean

    l = Len(fullName)
    If l = 0 Then fName = "": fPath = "": Exit Sub

    found = False
    For i = 1 To l
        c = Mid(fullName, l - i + 1, 1)
        If c = "\" Or c = ":" Then found = True: Exit For
    Next i

    If found Then
        fName = Right(fullName, i - 1)
        fPath = Left(fullName, l - i)
    Else
        fName = fullName
        fPath = ""
    End If
End Sub
' Strip NULL chars at the end of a string
Private Function stripNullChars(ByRef s As String) As String
Dim Pos As Long

    Pos = 0
    Do While Pos < Len(s)
        Pos = Pos + 1
        If (Asc(Mid(s, Pos, 1)) = 0) Then Exit Do
    Loop

    If (Pos > 0) And (Asc(Mid(s, Pos, 1)) = 0) Then
        stripNullChars = Left(s, Pos - 1)
    Else
        stripNullChars = s
    End If
End Function

'==================================================
#If False Then
Public Function GetLongPathName(ByVal sShortName As String) As String
    Dim sLongName As String
    Dim sTemp As String
    Dim iSlashPos As Integer
    If Len(sShortName) < 1 Then Exit Function
    If Right$(sShortName, 1) = "\" Then
        sShortName = Left$(sShortName, Len(sShortName) - 1)
    End If
    ' Add \ to short name to prevent Instr from failing
    sShortName = sShortName & "\"
    ' Start from 4 to ignore the "[Drive Letter]:\" characters
    If Mid$(sShortName, 2, 1) = ":" Then
        iSlashPos = 4
    ElseIf Left$(sShortName, 2) = "\\" Then     ' UNC path
        iSlashPos = InStr(3, sShortName, "\")   ' end of server name
        iSlashPos = InStr(iSlashPos + 1, sShortName, "\")   ' end of share name
        iSlashPos = iSlashPos + 1   ' start of first directory
    Else
        GetLongPathName = ""
        Exit Function
    End If
    iSlashPos = InStr(iSlashPos, sShortName, "\")
    ' Pull out each string between \ character for conversion
    While iSlashPos
        sTemp = Dir(Left$(sShortName, iSlashPos - 1), _
            vbNormal + vbHidden + vbSystem + vbDirectory)
        If sTemp = "" Then  ' Error 52 - Bad File Name or Number
            GetLongPathName = ""
            Exit Function
        End If
        sLongName = sLongName & "\" & sTemp
        iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
    Wend
    ' Prefix with the drive letter
    GetLongPathName = Left$(sShortName, 2) & sLongName
End Function
#End If

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetFileModTime
' Description:       Returns modification/creation time of a file
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       20/01/2004-12:36:06
'
' Parameters :       sFile (String)
'--------------------------------------------------------------------------------
'</CSCM>
Function GetFileModTime(sFile As String) As Date
    Dim dCreated As Date, dAccessed As Date, dModified As Date
    If GetFileTimes(sFile, dCreated, dAccessed, dModified, True) Then
        GetFileModTime = dModified
    Else
        GetFileModTime = Now
    End If
End Function

' Return false if there is an error.
Private Function GetFileTimes(ByVal file_name As String, _
    ByRef date_created As Date, ByRef date_accessed As _
    Date, ByRef date_written As Date, ByVal local_time As _
    Boolean) As Boolean
Dim file_handle As Long
Dim creation_time As FILETIME
Dim access_time As FILETIME
Dim write_time As FILETIME
Dim file_time As FILETIME

    GetFileTimes = False
    
    ' Open the file.
    file_handle = CreateFile(file_name, GENERIC_READ, _
        FILE_SHARE_READ Or FILE_SHARE_WRITE, _
        0&, OPEN_EXISTING, 0&, 0&)
    If file_handle = 0 Then
        Exit Function
    End If

    ' Get the times.
    If GetFileTime(file_handle, creation_time, _
        access_time, write_time) = 0 _
    Then
        Exit Function
    End If

    ' Close the file.
    If CloseHandle(file_handle) = 0 Then
        Exit Function
    End If

    ' See if we should convert to the local
    ' file system time.
    If local_time Then
        ' Convert to local file system time.
        FileTimeToLocalFileTime creation_time, file_time
        creation_time = file_time

        FileTimeToLocalFileTime access_time, file_time
        access_time = file_time

        FileTimeToLocalFileTime write_time, file_time
        write_time = file_time
    End If

    ' Convert into dates.
    date_created = FileTimeToDate(creation_time)
    date_accessed = FileTimeToDate(access_time)
    date_written = FileTimeToDate(write_time)

    GetFileTimes = True
End Function

' Convert the FILETIME structure into a Date.
Private Function FileTimeToDate(ft As FILETIME) As Date
' FILETIME units are 100s of nanoseconds.
Const TICKS_PER_SECOND = 10000000

Dim lo_time As Double
Dim hi_time As Double
Dim seconds As Double
Dim hours As Double
Dim the_date As Date

    ' Get the low order data.
    If ft.dwLowDateTime < 0 Then
        lo_time = 2 ^ 31 + (ft.dwLowDateTime And &H7FFFFFFF)
    Else
        lo_time = ft.dwLowDateTime
    End If

    ' Get the high order data.
    If ft.dwHighDateTime < 0 Then
        hi_time = 2 ^ 31 + (ft.dwHighDateTime And _
            &H7FFFFFFF)
    Else
        hi_time = ft.dwHighDateTime
    End If

    ' Combine them and turn the result into hours.
    seconds = (lo_time + 2 ^ 32 * hi_time) / _
        TICKS_PER_SECOND
    hours = CLng(seconds / 3600)
    seconds = seconds - hours * 3600

    ' Make the date.
    the_date = DateAdd("h", hours, "1/1/1601 0:00 AM")
    the_date = DateAdd("s", seconds, the_date)
    FileTimeToDate = the_date
End Function

