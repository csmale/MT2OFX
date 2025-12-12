Attribute VB_Name = "SpecialFolders"
Option Explicit

' $Header: /MT2OFX/SpecialFolders.bas 4     14/02/05 22:31 Colin $

Public Declare Function SHGetFolderPath Lib "shfolder.dll" Alias "SHGetFolderPathA" ( _
    ByVal hwndOwner As Long, _
    ByVal nFolder As Long, _
    ByVal hToken As Long, _
    ByVal dwFlags As Long, _
    ByVal sPath As String) As Long

Public Const CSIDL_DESKTOP = &H0                            ' <desktop>
Public Const CSIDL_INTERNET = &H1                           ' Internet Explorer (icon on desktop)
Public Const CSIDL_PROGRAMS = &H2                           ' Start Menu\Programs
Public Const CSIDL_CONTROLS = &H3                           ' My Computer\Control Panel
Public Const CSIDL_PRINTERS = &H4                           ' My Computer\Printers
Public Const CSIDL_PERSONAL = &H5                           ' My Documents
Public Const CSIDL_FAVORITES = &H6                          ' <user name>\Favorites
Public Const CSIDL_STARTUP = &H7                            ' Start Menu\Programs\Startup
Public Const CSIDL_RECENT = &H8                             ' <user name>\Recent
Public Const CSIDL_SENDTO = &H9                             ' <user name>\SendTo
Public Const CSIDL_BITBUCKET = &HA                          ' <desktop>\Recycle Bin
Public Const CSIDL_STARTMENU = &HB                          ' <user name>\Start Menu
Public Const CSIDL_MYDOCUMENTS = &HC                        ' logical "My Documents" desktop icon
Public Const CSIDL_MYMUSIC = &HD                            ' "My Music" folder
Public Const CSIDL_MYVIDEO = &HE                            ' "My Videos" folder
Public Const CSIDL_DESKTOPDIRECTORY = &H10                  ' <user name>\Desktop
Public Const CSIDL_DRIVES = &H11                            ' My Computer
Public Const CSIDL_NETWORK = &H12                           ' Network Neighborhood (My Network Places)
Public Const CSIDL_NETHOOD = &H13                           ' <user name>\nethood
Public Const CSIDL_FONTS = &H14                             ' windows\fonts
Public Const CSIDL_TEMPLATES = &H15
Public Const CSIDL_COMMON_STARTMENU = &H16                  ' All Users\Start Menu
Public Const CSIDL_COMMON_PROGRAMS = &H17                   ' All Users\Start Menu\Programs
Public Const CSIDL_COMMON_STARTUP = &H18                    ' All Users\Startup
Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19           ' All Users\Desktop
Public Const CSIDL_APPDATA = &H1A                           ' <user name>\Application Data
Public Const CSIDL_PRINTHOOD = &H1B                         ' <user name>\PrintHood
Public Const CSIDL_LOCAL_APPDATA = &H1C                     ' <user name>\Local Settings\Applicaiton Data (non roaming)
Public Const CSIDL_ALTSTARTUP = &H1D                        ' non localized startup
Public Const CSIDL_COMMON_ALTSTARTUP = &H1E                 ' non localized common startup
Public Const CSIDL_COMMON_FAVORITES = &H1F
Public Const CSIDL_INTERNET_CACHE = &H20
Public Const CSIDL_COOKIES = &H21
Public Const CSIDL_HISTORY = &H22
Public Const CSIDL_COMMON_APPDATA = &H23                    ' All Users\Application Data
Public Const CSIDL_WINDOWS = &H24                           ' GetWindowsDirectory()
Public Const CSIDL_SYSTEM = &H25                            ' GetSystemDirectory()
Public Const CSIDL_PROGRAM_FILES = &H26                     ' C:\Program Files
Public Const CSIDL_MYPICTURES = &H27                        ' C:\Program Files\My Pictures
Public Const CSIDL_PROFILE = &H28                           ' USERPROFILE
Public Const CSIDL_SYSTEMX86 = &H29                         ' x86 system directory on RISC
Public Const CSIDL_PROGRAM_FILESX86 = &H2A                  ' x86 C:\Program Files on RISC
Public Const CSIDL_PROGRAM_FILES_COMMON = &H2B              ' C:\Program Files\Common
Public Const CSIDL_PROGRAM_FILES_COMMONX86 = &H2C           ' x86 Program Files\Common on RISC
Public Const CSIDL_COMMON_TEMPLATES = &H2D                  ' All Users\Templates
Public Const CSIDL_COMMON_DOCUMENTS = &H2E                  ' All Users\Documents
Public Const CSIDL_COMMON_ADMINTOOLS = &H2F                 ' All Users\Start Menu\Programs\Administrative Tools
Public Const CSIDL_ADMINTOOLS = &H30                        ' <user name>\Start Menu\Programs\Administrative Tools
Public Const CSIDL_CONNECTIONS = &H31                       ' Network and Dial-up Connections
Public Const CSIDL_COMMON_MUSIC = &H35                      ' All Users\My Music
Public Const CSIDL_COMMON_PICTURES = &H36                   ' All Users\My Pictures
Public Const CSIDL_COMMON_VIDEO = &H37                      ' All Users\My Video
Public Const CSIDL_RESOURCES = &H38                         ' Resource Direcotry
Public Const CSIDL_RESOURCES_LOCALIZED = &H39               ' Localized Resource Direcotry
Public Const CSIDL_COMMON_OEM_LINKS = &H3A                  ' Links to All Users OEM specific apps
Public Const CSIDL_CDBURN_AREA = &H3B                       ' USERPROFILE\Local Settings\Application Data\Microsoft\CD Burning
' unused                               =&h003c
Public Const CSIDL_COMPUTERSNEARME = &H3D                   ' Computers Near Me (computered from Workgroup membership)

Public Const CSIDL_FLAG_CREATE = &H8000                     ' combine with CSIDL_ value to force folder creation in SHGetFolderPath()
Public Const CSIDL_FLAG_DONT_VERIFY = &H4000                ' combine with CSIDL_ value to return an unverified folder path
Public Const CSIDL_FLAG_NO_ALIAS = &H1000                   ' combine with CSIDL_ value to insure non-alias versions of the pidl
Public Const CSIDL_FLAG_PER_USER_INIT = &H800               ' combine with CSIDL_ value to indicate per-user init (eg. upgrade)
Public Const CSIDL_FLAG_MASK = &HFF00                       ' mask for all possible flag values

' for dwFlags:
Public Const SHGFP_TYPE_CURRENT = 0       ' current value for user, verify it exists
Public Const SHGFP_TYPE_DEFAULT = 1       ' default value, may not exist

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Public Type SHELLEXECUTEINFO
        cbSize As Long
        fMask As Long
        hWnd As Long
        lpVerb As String
        lpFile As String
        lpParameters As String
        lpDirectory As String
        nShow As Long
        hInstApp As Long
        '  Optional fields
        lpIDList As Long
        lpClass As String
        hkeyClass As Long
        dwHotKey As Long
        hIcon As Long
        hProcess As Long
End Type

Public Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" ( _
    ByRef lpExecInfo As SHELLEXECUTEINFO) As Long

Public Function GetSpecialFolder(iFolder As Long) As String
    Dim sTmp As String
    sTmp = String$(1024, vbNullChar)
    
    If SHGetFolderPath(0, iFolder, 0, SHGFP_TYPE_CURRENT, sTmp) = 0 Then
        GetSpecialFolder = Left$(sTmp, InStr(sTmp, vbNullChar) - 1)
    Else
        GetSpecialFolder = ""
    End If
End Function
