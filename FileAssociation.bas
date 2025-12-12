Attribute VB_Name = "FileAssociation"
Option Explicit

' $Header: /MT2OFX/FileAssociation.bas 8     15/06/09 19:24 Colin $

'// Windows Registry Messages
Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

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
Private Const ERROR_NO_MORE_ITEMS = 259

'// Windows Security Messages
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0

'// Windows Registry API calls
Private Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal hKey As Long) As Long

Private Declare Function RegCreateKeyEx _
 Lib "advapi32.dll" Alias "RegCreateKeyExA" _
(ByVal hKey As Long, _
 ByVal lpSubKey As String, _
 ByVal Reserved As Long, _
 ByVal lpClass As String, _
 ByVal dwOptions As Long, _
 ByVal samDesired As Long, _
 ByVal lpSecurityAttributes As Long, _
 phkResult As Long, _
 lpdwDisposition As Long) As Long

Private Declare Function RegOpenKeyEx _
 Lib "advapi32.dll" Alias "RegOpenKeyExA" _
(ByVal hKey As Long, _
 ByVal lpSubKey As String, _
 ByVal ulOptions As Long, _
 ByVal samDesired As Long, _
 phkResult As Long) As Long

Private Declare Function RegSetValueExString _
 Lib "advapi32.dll" Alias "RegSetValueExA" _
(ByVal hKey As Long, _
 ByVal lpValueName As String, _
 ByVal Reserved As Long, _
 ByVal dwType As Long, _
 ByVal lpValue As String, _
 ByVal cbData As Long) As Long

Private Declare Function RegSetValueExLong _
 Lib "advapi32.dll" Alias "RegSetValueExA" _
(ByVal hKey As Long, _
 ByVal lpValueName As String, _
 ByVal Reserved As Long, _
 ByVal dwType As Long, _
 lpValue As Long, _
 ByVal cbData As Long) As Long

Const FileType = "MT940File"

Private Sub CreateAssociation()

 Dim sPath As String
 sPath = Cfg.AppPath
 Dim sCommand As String
 sCommand = "Convert to OFX"
 
 '// File Associations begin with a listing
 '// of the default extension under HKEY_CLASSES_ROOT.
 '// So the first step is to create that
 '// root extension item
 CreateNewKey ".STA", HKEY_CLASSES_ROOT

 '// To the extension just added, add a
 '// subitem where the registry will look for
 '// commands relating to the .xxx extension
 '// ("MyApp.Document"). Its type is String (REG_SZ)
 SetKeyValue ".STA", "", FileType, REG_SZ

 '// Create the 'MyApp.Document' item under
 '// HKEY_CLASSES_ROOT. This is where you'll put
 '// the command line to execute or other shell
 '// statements necessary.
 CreateNewKey FileType & "\shell\" & sCommand & "\command", HKEY_CLASSES_ROOT

 '// Set its default item to "MyApp Document".
 '// This is what is displayed in Explorer against
 '// for files with a xxx extension. Its type is
 '// String (REG_SZ)
 SetKeyValue FileType, "", "SWIFT MT940 Bank Statement", REG_SZ

 '// Finally, add the path to myapp.exe
 '// Remember to add %1 as the final command
 '// parameter to assure the app opens the passed
 '// command line item.
 '// (results in '"c:\LongPathname\Myapp.exe %1")
 '// Again, its type is string.
 SetKeyValue FileType & "\shell\" & sCommand & "\command", "", sPath & " ""%1""", REG_SZ

 '// All done
    MyMsgBox "The file association has been made!"

End Sub

Private Function SetValueEx(ByVal hKey As Long, _
sValueName As String, lType As Long, _
vValue As Variant) As Long

 Dim nValue As Long
 Dim sValue As String

 Select Case lType
   Case REG_SZ
     sValue = vValue & vbNullChar
     SetValueEx = RegSetValueExString(hKey, _
       sValueName, 0&, lType, sValue, Len(sValue))

   Case REG_DWORD
     nValue = vValue
     SetValueEx = RegSetValueExLong(hKey, sValueName, _
       0&, lType, nValue, 4)

 End Select
  
End Function


Private Sub CreateNewKey(sNewKeyName As String, _
 lPredefinedKey As Long)

 '// handle to the new key
 Dim hKey As Long
 
 '// result of the RegCreateKeyEx function
 Dim r As Long
  
  r = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
   vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, r)

 Call RegCloseKey(hKey)

End Sub


Private Sub SetKeyValue(sKeyName As String, sValueName As String, _
vValueSetting As Variant, lValueType As Long)

 '// result of the SetValueEx function
 Dim r As Long
  
  '// handle of opened key
 Dim hKey As Long
  
  '// open the specified key
 r = RegOpenKeyEx(HKEY_CLASSES_ROOT, sKeyName, 0, _
   KEY_ALL_ACCESS, hKey)

 r = SetValueEx(hKey, sValueName, lValueType, vValueSetting)

 Call RegCloseKey(hKey)

End Sub

Private Sub cmdTest_Click()
 CreateAssociation
End Sub

