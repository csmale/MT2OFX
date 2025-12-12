Attribute VB_Name = "Registry"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : Registry
'    Project    : MT2OFX
'
'    Description: Registry access functions
'
'    Modified   : $Author: Colin $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/Registry.bas 4     15/06/09 19:25 Colin $"
' $History: Registry.bas $
' 
' *****************  Version 4  *****************
' User: Colin        Date: 15/06/09   Time: 19:25
' Updated in $/MT2OFX
' For transfer to new laptop
' 
' *****************  Version 4  *****************
' User: Colin        Date: 17/01/09   Time: 23:07
' Updated in $/MT2OFX
'
' *****************  Version 3  *****************
' User: Colin        Date: 25/11/08   Time: 22:23
' Updated in $/MT2OFX
' moving vss server!

'</CSCC>

'// Windows Registry Messages
Public Const REG_SZ As Long = 1
Public Const REG_DWORD As Long = 4
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

'// Windows Error Messages
Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_MORE_DATA = 234
Private Const ERROR_NO_MORE_ITEMS = 259

'// Windows Security Messages
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const KEY_ALL_ACCESS = &H3F
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const REG_OPTION_NON_VOLATILE = 0

'// Windows Registry API calls
Private Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal hKey As Long) As Long

Private Declare Function RegOpenKeyEx _
 Lib "advapi32.dll" Alias "RegOpenKeyExA" _
(ByVal hKey As Long, _
 ByVal lpSubKey As String, _
 ByVal ulOptions As Long, _
 ByVal samDesired As Long, _
 phkResult As Long) As Long

Private Declare Function RegQueryValueEx _
 Lib "advapi32.dll" Alias "RegQueryValueExA" _
 (ByVal hKey As Long, _
  ByVal lpValueName As String, _
  ByVal lpReserved As Long, _
  lpType As Long, _
  lpData As Any, _
  lpcbData As Long) As Long

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Declare Function RegEnumKeyEx _
    Lib "advapi32.dll" Alias "RegEnumKeyExA" _
    (ByVal hKey As Long, _
    ByVal dwIndex As Long, _
    ByVal lpName As String, _
    lpcbName As Long, _
    ByVal lpReserved As Long, _
    ByVal lpClass As String, _
    lpcbClass As Long, _
    lpftLastWriteTime As FILETIME) As Long

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetRegString
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       09/02/2005-12:48:39
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetRegString(hRoot As Long, sKey As String, sValue As String) As String
    Dim sVal As String
    Dim hKey As Long
    Dim r As Long
    Dim iKeyType As Long
    Dim iLen As Long
    
    r = RegOpenKeyEx(hRoot, sKey, 0, KEY_READ, hKey)
    If r <> 0 Then
        Debug.Print "(registry access error " & CStr(r) & " on " & sKey & ")"
        GetRegString = ""
    Else
        sVal = String$(1024, vbNullChar)
        iLen = Len(sVal)
        r = RegQueryValueEx(hKey, sValue, 0, iKeyType, ByVal sVal, iLen)
        If r <> 0 Then
            Debug.Print "(registry access error " & CStr(r) & " on " & sKey & ")"
            GetRegString = ""
        Else
' 20050411 CS: iLen can be zero if there is no value in the registry!!
            If iLen <= 0 Then
                GetRegString = ""
            Else
                GetRegString = Left$(sVal, iLen - 1)
            End If
        End If
        Call RegCloseKey(hKey)
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetRegStringRes
' Description:       As GetRegString but handles indirection through resource pointer
'                    in the form "@[path\]resource.dll,-resid"
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       1/11/2009-13:06:40
'
' Parameters :       hRoot (Long)
'                    sKey (String)
'                    sValue (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetRegStringRes(hRoot As Long, sKey As String, sValue As String) As String
    Dim sVal As String
    Dim iTmp As Long
    Dim iRes As Long
    Dim sTmp As String
    sVal = GetRegString(hRoot, sKey, sValue)
    If Left$(sVal, 1) = "@" Then
        iTmp = InStr(sVal, ",")
        If iTmp > 0 Then
            If IsNumeric(Mid$(sVal, iTmp + 1)) Then
                iRes = CLng(Mid$(sVal, iTmp + 1))
                sTmp = ExpandStrings(Mid$(sVal, 2, iTmp - 2))
                sTmp = LoadStringFromDll(sTmp, -iRes)
            End If
        End If
        sVal = sTmp
    End If
    GetRegStringRes = sVal
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetRegDword
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       15/02/2005-23:20:13
'
' Parameters :       hRoot (Long)
'                    sKey (String)
'                    sValue (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetRegDword(hRoot As Long, sKey As String, sValue As String) As Long
    Dim lVal As Long
    Dim hKey As Long
    Dim r As Long
    Dim iKeyType As Long
    Dim iLen As Long
    
    r = RegOpenKeyEx(hRoot, sKey, 0, KEY_READ, hKey)
    If r <> 0 Then
        Debug.Print "(registry access error " & CStr(r) & " on " & sKey & ")"
        GetRegDword = 0
    Else
        lVal = 0
        iLen = 4
        r = RegQueryValueEx(hKey, sValue, 0, iKeyType, lVal, iLen)
        If r <> 0 Then
            Debug.Print "(registry access error " & CStr(r) & " on " & sKey & ")"
            GetRegDword = 0
        Else
            GetRegDword = lVal
        End If
        Call RegCloseKey(hKey)
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       EnumRegKeys
' Description:       Return a list of subkeys as string array
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/17/2007-12:35:40
'
' Parameters :       hRoot (Long)
'                    sKey (String)
'                    asResult() (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function EnumRegKeys(hRoot As Long, sKey As String) As String()
    Dim hKey As Long
    Dim r As Long
    Dim iLen As Long
    Dim iIndex As Long
    Dim asTmp() As String
    Dim sTmp As String
    Dim sValue As String
    Dim ft As FILETIME
    
    ReDim asTmp(0 To 0)
    r = RegOpenKeyEx(hRoot, sKey, 0, KEY_READ, hKey)
    If r <> 0 Then
        Debug.Print "(registry access error " & CStr(r) & " on " & sKey & ")"
    Else
        iIndex = 0
        sValue = String$(1024, vbNullChar)
        iLen = Len(sValue)
        r = RegEnumKeyEx(hKey, iIndex, sValue, iLen, 0&, 0&, 0&, ft)
        Do While r = 0 Or r = ERROR_MORE_DATA
            asTmp(UBound(asTmp)) = Left$(sValue, iLen)
            ReDim Preserve asTmp(UBound(asTmp) + 1)
            iIndex = iIndex + 1
            sValue = String$(1024, vbNullChar)
            iLen = Len(sValue)
            r = RegEnumKeyEx(hKey, iIndex, sValue, iLen, 0&, 0&, 0&, ft)
        Loop
        If r <> ERROR_NO_MORE_ITEMS Then
            Debug.Print "(registry access error " & CStr(r) & " on " & sKey & ")"
        End If
        Call RegCloseKey(hKey)
    End If
    EnumRegKeys = asTmp
End Function


