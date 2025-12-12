Attribute VB_Name = "TempFiles"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : TempFiles
'    Project    : MT2OFX
'
'    Description:
'
'    Modified   : $Author: Colin $ $Date: 3/10/09 21:48 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/TempFiles.bas 9     3/10/09 21:48 Colin $"
' $History: TempFiles.bas $
' 
' *****************  Version 9  *****************
' User: Colin        Date: 3/10/09    Time: 21:48
' Updated in $/MT2OFX
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

Private Declare Function GetTempFileName Lib "kernel32" _
    Alias "GetTempFileNameA" (ByVal lpszPath As String, _
        ByVal lpPrefixString As String, _
        ByVal wUnique As Long, _
        ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" _
    Alias "GetTempPathA" (ByVal nBufferLength As Long, _
        ByVal lpBuffer As String) As Long

Private TempSeq As Long

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetTempFile
' Description:       Create a temporary file
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       09/10/2004-20:38:12
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetTempFile(Optional sExt As String = "TMP") As String
    Dim sTmp As String
    Dim sTempDir As String
    Dim lTmp As Long
    Dim iTempFile As Integer
    Dim sTempFile As String
    sTempDir = Space(1024)
    lTmp = GetTempPath(Len(sTempDir), sTempDir)
    If lTmp = 0 Then
        DBCSLog "", "GetTempPath failed"
        sTempDir = "C:\"
    Else
        DBCSLog sTempDir, "GetTempPath returned " & CStr(lTmp) & " chars"
'        sTmp = StrConv(sTempDir, vbFromUnicode)
'        sTmp = LeftB(sTmp, lChar)
'        sTempDir = StrConv(sTmp, vbUnicode)
        lTmp = InStr(sTempDir, vbNullChar) - 1
        sTempDir = Left$(sTempDir, lTmp)
        DBCSLog sTempDir, "Temp Dir now " & CStr(lTmp) & " chars"
    End If
    If Right$(sTempDir, 1) <> "\" Then
        sTempDir = sTempDir & "\"
    End If
    iTempFile = FreeFile
    On Error GoTo done
tryit:
    TempSeq = TempSeq + 1
    sTempFile = sTempDir & "MT" & Format(TempSeq, "00000") & "." & sExt
    Open sTempFile For Input Access Read As iTempFile
    ' if we are here the file exists so try again
    Close #iTempFile
    GoTo tryit
done:
    If Err.Number <> 53 Then    ' if not simply "file not found"...
        ShowError "GetTempFile", "Error creating temp file '" & sTempFile & "'"
        GetTempFile = ""
        Exit Function
    End If
    LogMessage False, True, "Created temp file " & sTempFile, ""
    DBCSLog sTempFile, "Created temp file"
    GetTempFile = sTempFile
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       RemoveTempFile
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       09/10/2004-22:10:06
'
' Parameters :       File (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub RemoveTempFile(File As String)
    If Len(File) = 0 Then
        Exit Sub
    End If
    LogMessage False, True, "Removing temp file " & File, ""
    On Error GoTo badkill
    Kill File
    Exit Sub
badkill:
    LogMessage False, True, "Error removing " & File & ": " & Err.Description
End Sub

