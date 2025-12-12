Attribute VB_Name = "SupportInfo"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : SupportInfo
'    Project    : MT2OFX
'
'    Description: Produces text report including information
'                 relevant to problem determination on a user's
'                 machine.
'
'    Modified   : $Author: Colin $ $Date: 15/11/10 0:31 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/SupportInfo.bas 16    15/11/10 0:31 Colin $"
' $History: SupportInfo.bas $
' 
' *****************  Version 16  *****************
' User: Colin        Date: 15/11/10   Time: 0:31
' Updated in $/MT2OFX
'
' *****************  Version 15  *****************
' User: Colin        Date: 24/11/09   Time: 22:05
' Updated in $/MT2OFX
' for 3.6 beta
'
' *****************  Version 14  *****************
' User: Colin        Date: 15/06/09   Time: 19:25
' Updated in $/MT2OFX
' For transfer to new laptop
'
' *****************  Version 14  *****************
' User: Colin        Date: 17/01/09   Time: 23:19
' Updated in $/MT2OFX
' Added identification of Windows 7 and 2008 Server (incl R2)
'
' *****************  Version 13  *****************
' User: Colin        Date: 25/11/08   Time: 22:24
' Updated in $/MT2OFX
' moving vss server!
'
' *****************  Version 11  *****************
' User: Colin        Date: 19/04/08   Time: 22:10
' Updated in $/MT2OFX
' move registry functions out
' add 64-bit detection
'
' *****************  Version 10  *****************
' User: Colin        Date: 7/12/06    Time: 14:46
' Updated in $/MT2OFX
' Updated for Windows Vista
'
' *****************  Version 8  *****************
' User: Colin        Date: 1/03/06    Time: 23:15
' Updated in $/MT2OFX
' Updated for latest versions of Money and Quicken
'
' *****************  Version 6  *****************
' User: Colin        Date: 11/04/05   Time: 23:31
' Updated in $/MT2OFX
' 3.3.4 bugfixes
'
' *****************  Version 5  *****************
' User: Colin        Date: 6/03/05    Time: 23:42
' Updated in $/MT2OFX
'</CSCC>


'*.vbs
'    version string?
'
'Last few lines of log file

Private oFile As TextStream
Private sFile As String
Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type
Public Type OSVERSIONINFOEX
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long    ' Enum OS_PLATFORM
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
        wServicePackMajor As Integer
        wServicePackMinor As Integer
        wSuiteMask As Integer   ' Enum OS_SUITE_MASK
        wProductType As Byte    ' Enum OS_PRODUCT_TYPE
        wReserved As Byte
End Type
Public Enum OS_SUITE_MASK
    VER_SUITE_SMALLBUSINESS = &H1
    VER_SUITE_ENTERPRISE = &H2
    VER_SUITE_BACKOFFICE = &H4
    VER_SUITE_TERMINAL = &H10
    VER_SUITE_SMALLBUSINESS_RESTRICTED = &H20
    VER_SUITE_EMBEDDEDNT = &H40
    VER_SUITE_DATACENTER = &H80
    VER_SUITE_SINGLEUSERTS = &H100
    VER_SUITE_PERSONAL = &H200
    VER_SUITE_BLADE = &H400
    VER_SUITE_STORAGE_SERVER = &H2000
    VER_SUITE_COMPUTE_SERVER = &H4000
End Enum
Public Enum OS_PRODUCT_TYPE
    VER_NT_WORKSTATION = 1
    VER_NT_DOMAIN_CONTROLLER = 2
    VER_NT_SERVER = 3
End Enum
Public Enum OS_PLATFORM
    VER_PLATFORM_WIN32s = 0
    VER_PLATFORM_WIN32_WINDOWS = 1
    VER_PLATFORM_WIN32_NT = 2
End Enum

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long

Public Const MAX_DEFAULTCHAR = 2
Public Const MAX_LEADBYTES = 12
Public Type CPINFO
        MaxCharSize As Long                    '  max length (Byte) of a char
        DefaultChar(1 To MAX_DEFAULTCHAR) As Byte   '  default character
        LeadByte(1 To MAX_LEADBYTES) As Byte        '  lead byte ranges
End Type
Public Type CPINFOEX
        MaxCharSize As Long                    '  max length (Byte) of a char
        DefaultChar(1 To MAX_DEFAULTCHAR) As Byte   '  default character
        LeadByte(1 To MAX_LEADBYTES) As Byte        '  lead byte ranges
        UnicodeDefaultChar As Integer          '
        CodePage As Long                 ' you read it right - this is a STRING
        CodePageName As String * 260            ' MAX_PATH
End Type

Public Declare Function GetACP Lib "kernel32" () As Long
Public Declare Function GetOEMCP Lib "kernel32" () As Long
Public Declare Function GetCPInfo Lib "kernel32" (ByVal CodePage As Long, lpCPInfo As CPINFO) As Long
Public Declare Function GetCPInfoEx Lib "kernel32" Alias "GetCPInfoExA" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpCPInfoEx As CPINFOEX) As Long
Private Type SYSTEM_INFO
        wProcessorArchitecture As Integer
        wReserved As Integer
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type

Private Declare Sub GetNativeSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Const PROCESSOR_ARCHITECTURE_IA64 = 6
Private Const PROCESSOR_ARCHITECTURE_AMD64 = 9
Private Const SM_SERVERR2 = 89
Private Declare Function GetSystemMetrics Lib "User32" (ByVal nIndex As Long) As Long

' Vista-specific stuff
' NB: GetProductInfo is only implemented on Vista/Longhorn
' NB: useful return value is in 5th parameter!
Private Declare Function GetProductInfo Lib "kernel32" ( _
        ByVal dwOSMajorVersion As Long, _
        ByVal dwOSMinorVersion As Long, _
        ByVal dwSpMajorVersion As Long, _
        ByVal dwSpMinorVersion As Long, _
        ByRef pdwReturnedProductType As Long _
        ) As Long
Private Const PRODUCT_DATACENTER_SERVER As Long = &H8               ' Windows Server 2008, Datacenter Edition
Private Const PRODUCT_DATACENTER_SERVER_CORE As Long = &HC          ' Windows Server 2008, Datacenter Edition (Server Core installation)
Private Const PRODUCT_DATACENTER_SERVER_V As Long = &H25            ' Windows Server 2008, Datacenter Edition, no Hyper-V
Private Const PRODUCT_DATACENTER_SERVER_CORE_V As Long = &H27       ' Windows Server 2008, Datacenter Edition (Server Core installation), no Hyper-V
Private Const PRODUCT_ENTERPRISE_SERVER As Long = &HA               ' Windows Server 2008, Enterprise Edition
Private Const PRODUCT_ENTERPRISE_SERVER_CORE As Long = &HE          ' Windows Server 2008, Enterprise Edition (Server Core installation)
Private Const PRODUCT_ENTERPRISE_SERVER_IA64 As Long = &HF          ' Windows Server 2008, Enterprise Edition for Itanium-based Systems
Private Const PRODUCT_HOME_BASIC As Long = &H2
Private Const PRODUCT_HOME_PREMIUM As Long = &H3
Private Const PRODUCT_PRO_PREMIUM As Long = &H4
Private Const PRODUCT_PRO_SMALLBUSINESS As Long = &H5
Private Const PRODUCT_PRO_STANDARD As Long = &H6
Private Const PRODUCT_SMALLBUSINESS_SERVER As Long = &H9
Private Const PRODUCT_STANDARD_SERVER As Long = &H7                 ' Windows Server 2008
Private Const PRODUCT_STANDARD_SERVER_CORE As Long = &HD            ' Windows Server 2008 (Server Core installation)
Private Const PRODUCT_STARTER As Long = &HB
Private Const PRODUCT_ULTIMATE As Long = &H1
Private Const PRODUCT_ULTIMATE_N As Long = &H47
Private Const PRODUCT_ULTIMATE_E As Long = &H1C
Private Const PRODUCT_BUSINESS_N As Long = &H10
Private Const PRODUCT_CLUSTER_SERVER As Long = &H12
Private Const PRODUCT_ENTERPRISE As Long = &H4
Private Const PRODUCT_ENTERPRISE_E As Long = &H46
Private Const PRODUCT_ENTERPRISE_N As Long = &H1B






'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoSupportInfo
' Description:       prepare support log file
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       08/02/2005-21:23:14
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub DoSupportInfo()
    Dim oFSO As New FileSystemObject
' 20061129 CS: changed to AppData for Vista compatibility
    sFile = Cfg.AppDataPath & "\support.txt"
    Set oFile = oFSO.OpenTextFile(sFile, ForWriting, True, TristateFalse)
    DoFileHeader sFile
    DoProgramInfo
    DoSystemInfo
    DoLocaleInfo LOCALE_SYSTEM_DEFAULT, "System Locale", GetSystemDefaultLangID()
    DoLocaleInfo LOCALE_USER_DEFAULT, "User Locale", GetUserDefaultLangID()
    DoDllInfoAll
    DoInstalledPrograms
    DoIniFileInfo
    DoScriptFileInfo
    DoFiletypesAll
    DoLogFileTail Cfg.SptLogFileLines
    oFile.Close
    Set oFile = Nothing
    DoDefaultFileAction sFile
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoIniFileInfo
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       09/02/2005-00:12:49
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DoIniFileInfo()
    Dim sIni As String
    Dim sLine As String
    Dim iFile2 As Integer
    iFile2 = FreeFile
    oFile.WriteLine "Contents of MT2OFX.INI"
    oFile.WriteLine "======================"
' 20061129 CS: changed for Vista compatibility
    sIni = Cfg.IniFileName
    
    Open sIni For Input Access Read As iFile2
    
    Do While Not EOF(iFile2)
        Line Input #iFile2, sLine
        oFile.WriteLine sLine
    Loop
    
    Close #iFile2
    oFile.WriteLine "=== end of MT2OFX.INI ==="
    oFile.WriteLine ""
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoScriptFileInfo
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       09/02/2005-00:13:16
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DoScriptFileInfo()
    Dim sFile As String
    Dim dMod As Date
    Dim sPath As String
    Dim lBytes As Long
    
    oFile.WriteLine "Scripts in application directory"
    oFile.WriteLine "================================"
    sFile = Dir(Cfg.ScriptPath & "\*.vbs")
    Do While sFile <> ""
        sPath = Cfg.ScriptPath & "\" & sFile
        dMod = GetFileModTime(sPath)
        lBytes = FileLen(sPath)
        oFile.WriteLine sFile & vbTab & lBytes & vbTab & FormatSupportDate(dMod)
        sFile = Dir
    Loop
    oFile.WriteLine ""
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoLocaleInfo
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       08/02/2005-23:22:06
'
' Parameters :       dwWhat (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DoLocaleInfo(dwWhat As Long, sTitle As String, lLCID As Long)
    Dim sTmp As String
    Dim sTmp2 As String
    oFile.WriteLine sTitle
    oFile.WriteLine String$(Len(sTitle), "=")
    oFile.WriteLine "Locale: " & GetLocaleString(dwWhat, LOCALE_SLANGUAGE, "")
    oFile.WriteLine "LCID: &H" & Hex(lLCID)
    sTmp = GetLocaleString(dwWhat, LOCALE_SENGCOUNTRY, "")
    sTmp2 = GetLocaleString(dwWhat, LOCALE_SISO3166CTRYNAME, "")
    oFile.WriteLine "Country: " & sTmp & " (" & sTmp2 & ")"
    sTmp = GetLocaleString(dwWhat, LOCALE_SENGLANGUAGE, "")
    sTmp2 = GetLocaleString(dwWhat, LOCALE_SISO639LANGNAME, "")
    oFile.WriteLine "Language: " & sTmp & " (" & sTmp2 & ")"
    sTmp = GetLocaleString(dwWhat, LOCALE_SENGCURRNAME, "")
    sTmp2 = GetLocaleString(dwWhat, LOCALE_SINTLSYMBOL, "")
    oFile.WriteLine "Currency: " & sTmp & " (" & sTmp2 & ")"
    sTmp = GetLocaleString(dwWhat, LOCALE_SDECIMAL, "")
    oFile.WriteLine "Decimal separator: " & sTmp
    sTmp = GetLocaleString(dwWhat, LOCALE_STHOUSAND, "")
    oFile.WriteLine "Thousands separator: " & sTmp
    sTmp = GetLocaleString(dwWhat, LOCALE_SLIST, "")
    oFile.WriteLine "List separator: " & sTmp
    sTmp = GetLocaleString(dwWhat, LOCALE_SSHORTDATE, "")
    oFile.WriteLine "Short date format: " & sTmp
    sTmp = GetLocaleString(dwWhat, LOCALE_SLONGDATE, "")
    oFile.WriteLine "Long date format: " & sTmp
    oFile.WriteLine ""
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoCodePage
' Description:       Print Code Page information
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       8/20/2006-00:07:24
'
' Parameters :       sText (String)
'                    cp (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DoCodePage(sText As String, cp As Long)
    Dim cpi As CPINFOEX
    Dim iTmp As Long
    cpi.CodePageName = ""
    If GetCPInfoEx(cp, 0, cpi) = 0 Then
        oFile.WriteLine "Unable to get information for " & sText & " code page " & CStr(cp)
    Else
        iTmp = InStr(cpi.CodePageName, vbNullChar)
        oFile.WriteLine sText & " Code Page: " & CStr(cpi.CodePage) & " = " & Left$(cpi.CodePageName, iTmp - 1)
        oFile.WriteLine "Double-byte: " & IIf(cpi.MaxCharSize > 1, "Yes", "No")
    End If
    oFile.WriteLine ""
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoSystemInfo
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       08/02/2005-22:13:37
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DoSystemInfo()
    Dim osvi_old As OSVERSIONINFO   ' only used to get the size
    Dim osvi As OSVERSIONINFOEX
    
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        osvi.dwOSVersionInfoSize = Len(osvi_old)
        If GetVersionEx(osvi) = 0 Then
            oFile.WriteLine "Windows version not available, err=" & GetLastError()
        Else
            DumpVersionInfo osvi
        End If
    Else
        DumpVersionInfo osvi
    End If
    oFile.WriteBlankLines 1
    DoCodePage "ANSI", GetACP()
    DoCodePage "OEM", GetOEMCP()
End Sub
Private Sub DumpVersionInfo(osvi As OSVERSIONINFOEX)
    Dim sTmp As String
    
    oFile.WriteLine GetSystemString(osvi)
    sTmp = "Windows version " & CStr(osvi.dwMajorVersion) & "." & CStr(osvi.dwMinorVersion) _
        & " " & Left$(osvi.szCSDVersion, InStr(osvi.szCSDVersion, vbNullChar) - 1)
    oFile.WriteLine sTmp
    
    If osvi.dwOSVersionInfoSize = Len(osvi) Then
        sTmp = "Service pack " & CStr(osvi.wServicePackMajor) & "." & CStr(osvi.wServicePackMinor) _
            & ", Suite &H" & Hex(osvi.wSuiteMask) & ", Product Type &H" & Hex(osvi.wProductType)
        oFile.WriteLine sTmp
    End If
End Sub
Private Function GetSystemString(osvi As OSVERSIONINFOEX) As String
    Dim sTmp As String, sTmp2 As String
    Dim si As SYSTEM_INFO
    Dim iProd As Long
    sTmp = ""
    Select Case osvi.dwPlatformId
    Case VER_PLATFORM_WIN32_NT
        If osvi.dwMajorVersion = 6 And (osvi.dwMinorVersion = 0 Or osvi.dwMinorVersion = 1) Then
            If osvi.wProductType = VER_NT_WORKSTATION Then
                If osvi.dwMinorVersion = 0 Then
                    sTmp = "Microsoft Windows Vista"
                Else
                    sTmp = "Microsoft Windows 7"
                End If
            Else
                sTmp = "Windows Server 2008"
                If osvi.dwMinorVersion = 1 Then
                    sTmp = sTmp & " R2"
                End If
            End If
' NB: GetProductInfo is Vista/Longhorn only!
            If GetProductInfo(osvi.dwMajorVersion, osvi.dwMinorVersion, _
                            osvi.wServicePackMajor, osvi.wServicePackMinor, iProd) = 0 Then
                sTmp2 = "(unknown or invalid product type, or invalid license)"
            Else
                Select Case iProd
                Case PRODUCT_DATACENTER_SERVER:         sTmp2 = "Datacentre Edition"
                Case PRODUCT_DATACENTER_SERVER_CORE:    sTmp2 = "Datacentre Edition (Server Core installation)"
                Case PRODUCT_ENTERPRISE_SERVER:         sTmp2 = "Enterprise Edition"
                Case PRODUCT_ENTERPRISE_SERVER_CORE:    sTmp2 = "Enterprise Edition (Server Core installation)"
                Case PRODUCT_ENTERPRISE_SERVER_IA64:    sTmp2 = "Enterprise Edition for Itanium-based Systems"
                Case PRODUCT_HOME_BASIC:                sTmp2 = "Home Basic Edition"
                Case PRODUCT_HOME_PREMIUM:              sTmp2 = "Home Premium Edition"
                Case PRODUCT_PRO_PREMIUM:               sTmp2 = "Professional Premium Edition"
                Case PRODUCT_PRO_SMALLBUSINESS:         sTmp2 = "Professional Small Business Edition"
                Case PRODUCT_PRO_STANDARD:              sTmp2 = "Professional Standard Edition"
                Case PRODUCT_SMALLBUSINESS_SERVER:      sTmp2 = "Small Business Server"
                Case PRODUCT_STANDARD_SERVER:           sTmp2 = "Standard Edition"
                Case PRODUCT_STANDARD_SERVER_CORE:      sTmp2 = "(Server Core installation)"
                Case PRODUCT_STARTER:                   sTmp2 = "Starter Edition"
                Case PRODUCT_ULTIMATE:                  sTmp2 = "Ultimate Edition"
                Case PRODUCT_ULTIMATE_E:                sTmp2 = "Ultimate Edition E"
                Case PRODUCT_ULTIMATE_N:                sTmp2 = "Ultimate Edition N"
                Case Else:                              sTmp2 = "Product &H" & Hex(iProd)
                End Select
            End If
            sTmp = sTmp & ", " & sTmp2
            GetNativeSystemInfo si
            If si.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64 Then
                sTmp = sTmp & " (64-bit)"
            End If
        ElseIf osvi.dwMajorVersion = 5 And osvi.dwMinorVersion = 2 Then
            GetNativeSystemInfo si
            If GetSystemMetrics(SM_SERVERR2) > 0 Then
                sTmp = "Microsoft Windows Server 2003 R2"
            ElseIf si.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_IA64 And osvi.wProductType = VER_NT_WORKSTATION Then
                sTmp = "Microsoft Windows XP Professional x64 Edition"
            Else
                sTmp = "Microsoft Windows Server 2003"
            End If
        ElseIf osvi.dwMajorVersion = 5 And osvi.dwMinorVersion = 1 Then
            sTmp = "Microsoft Windows XP"
        ElseIf osvi.dwMajorVersion = 5 And osvi.dwMinorVersion = 0 Then
            sTmp = "Microsoft Windows 2000"
        ElseIf osvi.dwMajorVersion <= 4 Then
            sTmp = "Microsoft Windows NT"
        End If
        If osvi.dwOSVersionInfoSize = Len(osvi) Then
        ' Test for the workstation type.
            If osvi.wProductType = VER_NT_WORKSTATION Then
                If osvi.dwMajorVersion = 4 Then
                    sTmp = sTmp & " Workstation 4.0"
                ElseIf osvi.dwMajorVersion = 5 Then
                    If (osvi.wSuiteMask And VER_SUITE_PERSONAL) <> 0 Then
                        sTmp = sTmp & " Home Edition"
                    Else
                        sTmp = sTmp & " Professional"
                    End If
                End If
        ' Test for the server type.
            ElseIf (osvi.wProductType = VER_NT_SERVER) Or (osvi.wProductType = VER_NT_DOMAIN_CONTROLLER) Then
                If osvi.dwMajorVersion = 5 And osvi.dwMinorVersion = 2 Then
                    If si.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_IA64 Then
                        If (osvi.wSuiteMask And VER_SUITE_DATACENTER) <> 0 Then
                            sTmp = sTmp & " Datacenter Edition for Itanium-based Systems"
                        ElseIf (osvi.wSuiteMask & VER_SUITE_ENTERPRISE) <> 0 Then
                            sTmp = sTmp & " Enterprise Edition for Itanium-based Systems"
                        End If
                    ElseIf si.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64 Then
                        If (osvi.wSuiteMask And VER_SUITE_DATACENTER) <> 0 Then
                            sTmp = sTmp & " Datacenter x64 Edition"
                        ElseIf (osvi.wSuiteMask & VER_SUITE_ENTERPRISE) <> 0 Then
                            sTmp = sTmp & " Enterprise x64 Edition"
                        Else
                            sTmp = sTmp & " Standard x64 Edition"
                        End If
                    Else
                        If (osvi.wSuiteMask & VER_SUITE_DATACENTER) <> 0 Then
                            sTmp = sTmp & " Datacenter Edition"
                        ElseIf (osvi.wSuiteMask & VER_SUITE_ENTERPRISE) <> 0 Then
                            sTmp = sTmp & " Enterprise Edition"
                        ElseIf (osvi.wSuiteMask = VER_SUITE_BLADE) <> 0 Then
                            sTmp = sTmp & "Web Edition"
                        Else
                            sTmp = sTmp & "Standard Edition"
                        End If
                    End If
                ElseIf (osvi.dwMajorVersion = 5) And (osvi.dwMinorVersion = 0) Then
                    If (osvi.wSuiteMask & VER_SUITE_DATACENTER) <> 0 Then
                        sTmp = sTmp & " Datacenter Server"
                    ElseIf (osvi.wSuiteMask & VER_SUITE_ENTERPRISE) <> 0 Then
                        sTmp = sTmp & " Advanced Server"
                    Else
                        sTmp = sTmp & " Server"
                    End If
                Else  ' Windows NT 4.0
                    If (osvi.wSuiteMask & VER_SUITE_ENTERPRISE) <> 0 Then
                       sTmp = sTmp & "Server 4.0, Enterprise Edition"
                    Else
                        sTmp = sTmp & "Server 4.0"
                    End If
                End If
            End If
        End If
    Case VER_PLATFORM_WIN32_WINDOWS
        If osvi.dwMajorVersion = 4 And osvi.dwMinorVersion = 0 Then
            sTmp = "Windows 95"
            If Mid$(osvi.szCSDVersion, 2, 1) = "B" Or Mid$(osvi.szCSDVersion, 2, 1) = "C" Then
                sTmp = sTmp & " OSR2"
            End If
        ElseIf osvi.dwMajorVersion = 4 And osvi.dwMinorVersion = 10 Then
            sTmp = "Windows 98"
            If Mid$(osvi.szCSDVersion, 2, 1) = "A" Or Mid$(osvi.szCSDVersion, 2, 1) = "B" Then
                sTmp = sTmp & " SE"
            End If
        ElseIf osvi.dwMajorVersion = 4 And osvi.dwMinorVersion = 90 Then
            sTmp = "Windows Millennium Edition"
        Else
        End If
    Case VER_PLATFORM_WIN32s
        sTmp = "Win32s"
    End Select
    If Len(sTmp) = 0 Then
        sTmp = "Unknown Windows version " & CStr(osvi.dwMajorVersion) & "." & CStr(osvi.dwMinorVersion)
    End If
    GetSystemString = sTmp
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoProgramInfo
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       08/02/2005-22:04:14
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DoProgramInfo()
    Dim sPath As String
    sPath = App.Path & "\" & App.EXEName & ".exe"
    Dim sTmp As String
    sTmp = "MT2OFX Program version: " & CStr(App.Major) _
        & "." & CStr(App.Minor) & "." & CStr(App.Revision) _
        & vbTab & FileLen(sPath) _
        & vbTab & FormatSupportDateTime(GetFileModTime(sPath))
    oFile.WriteLine sTmp
    oFile.WriteLine "Running " & sPath
    oFile.WriteLine "MT2OFX.INI  : " & Cfg.IniFileName
    oFile.WriteLine "App path    : " & Cfg.AppPath
    oFile.WriteLine "My documents: " & Cfg.MyDocumentsPath
    oFile.WriteLine "User data   : " & Cfg.AppDataPath
    oFile.WriteLine "Resources   : " & Cfg.ResourcePath
    oFile.WriteLine "Scripts     : " & Cfg.ScriptPath
    oFile.WriteLine ""
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoDllInfoAll
' Description:       log info on all DLLs
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       08/02/2005-21:28:36
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DoDllInfoAll()
    DoDllInfo "USER32.DLL"
    DoDllInfo "KERNEL32.DLL"
    DoDllInfo "OLE32.DLL"
' 20060823 CS: Added unicows for unicode support on win98 etc
    DoDllInfo "UNICOWS.DLL"
    DoDllInfo "OLEAUT32.DLL"
    DoDllInfo "ADVAPI32.DLL"
    DoDllInfo "MSVCRT.DLL"
    DoDllInfo "MSVBVM60.DLL"
    DoDllInfo "COMDLG32.OCX"
    DoDllInfo "COMDLG32.DLL"
    DoDllInfo "MSSCRIPT.OCX"
    DoDllInfo "MSCOMCTL.OCX"
    DoDllInfo "COMCTL32.DLL"
    DoDllInfo "MSCOMCT2.OCX"
    DoDllInfo "WSHEXT.DLL"
    DoDllInfo "SCRRUN.DLL"
    DoDllInfo "VBSCRIPT.DLL"
    DoDllInfo "MSXML3.DLL"
    DoDllInfo "FM20.DLL"
    DoDllInfo "CRYPT32.DLL"
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoDllInfo
' Description:       log info on one dll
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       08/02/2005-21:29:03
'
' Parameters :       sDll (String)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DoDllInfo(sDLL As String)
    Dim sTmp As String
    Dim sVerFile As String
    Dim sVerMin As String
    Dim dllInfo As VS_FIXEDFILEINFO
    Dim dllStrings As FILEINFO
    Dim dFileDate As Date, sFileDate As String
    
    sTmp = GetModulePath(sDLL)
    If Len(sTmp) = 0 Then
        sTmp = sDLL
        sFileDate = "Unknown"
    Else
        dFileDate = GetFileModTime(sTmp)
        sFileDate = FormatSupportDateTime(dFileDate)
    End If
    oFile.WriteLine "Information for " & sTmp
    oFile.WriteLine "Modification date: " & sFileDate
    If GetModuleInfo(sDLL, dllInfo) Then
        sVerFile = GetVersionString(dllInfo.dwFileVersionMS, dllInfo.dwFileVersionLS)
        oFile.WriteLine "Version: " & sVerFile
        If GetFileVersionInformation(sDLL, dllStrings) = eOK Then
            oFile.WriteLine "Language: " & dllStrings.Language
        End If
    Else
        oFile.WriteLine "No module information available"
    End If
    oFile.WriteLine ""
End Sub


'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoFiletype
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       09/02/2005-12:30:20
'
' Parameters :       sType (String)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DoFiletype(sType As String)
    Dim sProgClass As String
    Dim sFileType As String
    Dim sDefVerb As String
    Dim sAction As String
    Dim sTmp As String
    Dim sDDECommand As String
    Dim sDDEApp As String
    Dim sDDETopic As String
    Dim hKey As Long
    Dim r As Long
    Dim iKeyType As Long
    Dim iLen As Long
    
    oFile.WriteLine "File association information for " & sType
' from Vista onwards associations are handled a bit differently
    sProgClass = GetRegString(HKEY_CURRENT_USER, "\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & sType & "\UserChoice", "Progid")
    If Len(sProgClass) = 0 Then
        sProgClass = GetRegString(HKEY_CLASSES_ROOT, sType, "")
    End If
    If Len(sProgClass) > 0 Then
        oFile.WriteLine "Program class: " & sProgClass
        sFileType = GetRegString(HKEY_CLASSES_ROOT, sProgClass, "")
        oFile.WriteLine "File Type: " & sFileType
        sDefVerb = GetRegString(HKEY_CLASSES_ROOT, sProgClass & "\shell", "")
        oFile.WriteLine "Default verb: " & sDefVerb
        If sDefVerb = "" Then
            sDefVerb = "open"
        End If
        sTmp = sProgClass & "\shell\" & sDefVerb
        sAction = GetRegString(HKEY_CLASSES_ROOT, sTmp & "\command", "")
        oFile.WriteLine "Default verb action: " & sAction
        sDDECommand = GetRegString(HKEY_CLASSES_ROOT, sTmp & "\ddeexec", "")
        If sDDECommand <> "" Then
            sDDEApp = GetRegString(HKEY_CLASSES_ROOT, sTmp & "\ddeexec\application", "")
            sDDETopic = GetRegString(HKEY_CLASSES_ROOT, sTmp & "\ddeexec\topic", "")
            oFile.WriteLine "DDE command: " & sDDECommand & " app=" & sDDEApp & " topic=" & sDDETopic
        End If
    Else
        oFile.WriteLine "Extension not registered"
    End If
    oFile.WriteLine ""
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoFiletypesAll
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       09/02/2005-12:30:03
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DoFiletypesAll()
    oFile.WriteLine "File Type Associations"
    oFile.WriteLine "======================"
    DoFiletype ".sta"
    DoFiletype ".940"
    DoFiletype ".swi"
    DoFiletype ".mt"
    DoFiletype ".txt"
    DoFiletype ".csv"
    DoFiletype ".ofx"
    DoFiletype ".ofc"
    DoFiletype ".qif"
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoLogFileTail
' Description:       copy the last 'n' lines from the log file
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       09/02/2005-22:22:06
'
' Parameters :       iLines (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DoLogFileTail(iLines As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim sLines() As String
    ReDim sLines(1 To iLines)
    Dim iLogFile As Integer
    Dim bWrapped As Boolean
    Dim sLine As String
    Dim sTmp As String
    
    sTmp = "Log File (last " & CStr(iLines) & " lines)"
    oFile.WriteLine sTmp
    oFile.WriteLine String$(Len(sTmp), "=")
    
    If Cfg.LogFile = "" Then
        oFile.WriteLine "No log file configured"
        oFile.WriteLine ""
        Exit Sub
    End If
    
    CloseLogFile
    iLogFile = FreeFile
    If PathIsRelative(Cfg.LogFile) Then
' 20061129 CS: changed to AppData for Vista compatibility
        sTmp = Cfg.AppDataPath & "\" & Cfg.LogFile
    Else
        sTmp = Cfg.LogFile
    End If
    oFile.WriteLine "Log file: " & sTmp
' run thru entire file, remembering last few lines in the ring buffer
    On Error GoTo nofile
    Open sTmp For Input Access Read As iLogFile
    On Error Resume Next
    i = 0
    bWrapped = False
    Do While Not EOF(iLogFile)
        Line Input #iLogFile, sLine
        i = i + 1
        If i > iLines Then
            i = 1
            bWrapped = True
        End If
        sLines(i) = sLine
    Loop
    Close #iLogFile
' now copy out the ring buffer to the support log
    If bWrapped Then
        For j = i + 1 To iLines
            oFile.WriteLine sLines(j)
        Next
    End If
    For j = 1 To i
        oFile.WriteLine sLines(j)
    Next
goback:
    oFile.WriteLine ""
    Exit Sub
nofile:
    oFile.WriteLine "Error opening log file: " & Err.Description
    Resume goback
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoFileHeader
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       14/02/2005-15:40:22
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DoFileHeader(sFile As String)
    Dim dNow As Date
    dNow = Now()
    oFile.WriteLine "This file contains a summary of important information about your version of MT2OFX"
    oFile.WriteLine "and its environment. If you report a fault, or ask for support, please enclose this"
    oFile.WriteLine "file in full. This will often enable me to pinpoint the problem without needing to"
    oFile.WriteLine "ask for technical details about your system."
    oFile.WriteLine ""
    oFile.WriteLine "Please review the file for confidential information before sending it."
    oFile.WriteLine ""
    oFile.WriteLine "The file was created on: " & FormatSupportDateTime(dNow)
    oFile.WriteLine "This file is stored in: " & sFile
    oFile.WriteLine ""
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoInstalledPrograms
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       15/02/2005-23:08:07
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DoInstalledPrograms()
    oFile.WriteLine "Installed programs"
    oFile.WriteLine "=================="
    DoMoney
    DoQuicken
    DoAceMoney
    DoMoneyDance
    oFile.WriteLine ""
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoMoney
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       15/02/2005-23:08:23
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DoMoney()
    Dim iVer As Integer
    Dim sVer As String
    Dim sPath As String
    Dim sReg As String
    Dim sSKU As String
    Dim sFile As String
    Dim sCountry As String
    Dim lLCID As Long
    Dim sVerFile As String
    Dim dllInfo As VS_FIXEDFILEINFO
    Dim dllStrings As FILEINFO
    
' old style - works at least with 2003 (v11) but no longer with 2005 (v14)
    For iVer = 7 To 13
        sVer = CStr(iVer) & ".0"
        sReg = "Software\Microsoft\Money\" & sVer
        sPath = GetRegString(HKEY_LOCAL_MACHINE, sReg, "MoneyPath")
        If Len(sPath) > 0 Then
            sSKU = GetRegString(HKEY_LOCAL_MACHINE, sReg, "SKU")
            oFile.WriteLine "Microsoft Money V" & sVer & " " & sSKU & " found at " & sPath
            lLCID = GetRegDword(HKEY_LOCAL_MACHINE, sReg, "LCID")
            oFile.WriteLine "LCID=&H" & Hex(lLCID) & " (" & CStr(lLCID) & ")"
            sFile = GetRegString(HKEY_LOCAL_MACHINE, sReg, "ExePath")
            If sFile = "" Then
                sFile = GetRegString(HKEY_LOCAL_MACHINE, sReg, "MoneyPath") & "\msmoney.exe"
            End If
            If Dir(sFile) = "" Then
                sFile = GetRegString(HKEY_LOCAL_MACHINE, sReg, "MoneyPath") & "\system\msmoney.exe"
            End If
            If Dir(sFile) <> "" Then
                If GetModuleInfo(sFile, dllInfo) Then
                    sVerFile = GetVersionString(dllInfo.dwFileVersionMS, dllInfo.dwFileVersionLS)
                    oFile.WriteLine "Version: " & sVerFile
                    If GetFileVersionInformation(sFile, dllStrings) = eOK Then
                        oFile.WriteLine "Language: " & dllStrings.Language
                    End If
                Else
                    oFile.WriteLine "No version information available"
                End If
            Else
                oFile.WriteLine sFile & " not found!!"
            End If
            oFile.WriteLine ""
        End If
    Next
' new style - started somewhere after 2003 (v11)
    For iVer = 12 To 20
        sVer = CStr(iVer) & ".0"
        sReg = "Software\Microsoft\Money\" & sVer
        sPath = GetRegString(HKEY_LOCAL_MACHINE, sReg & "\Setup", "InstallDir")
        If Len(sPath) > 0 Then
            oFile.WriteLine "Microsoft Money V" & sVer & " found at " & sPath
            sCountry = GetRegString(HKEY_LOCAL_MACHINE, sReg & "\Setup", "CountryID")
            oFile.WriteLine "Country ID: " & sCountry
            sFile = sPath & "\msmoney.exe"
            If Dir(sFile) <> "" Then
                If GetModuleInfo(sFile, dllInfo) Then
                    sVerFile = GetVersionString(dllInfo.dwFileVersionMS, dllInfo.dwFileVersionLS)
                    oFile.WriteLine "Version: " & sVerFile
                Else
                    oFile.WriteLine "No version information available"
                End If
            Else
                oFile.WriteLine sFile & " not found!!"
            End If
            oFile.WriteLine ""
        End If
    Next
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoQuicken
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       15/02/2005-23:53:22
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DoQuicken()
    Dim sQuickenProg As String
    Dim sQuickenVer As String
    Dim sTmp As String
    
' 20051221 CS: Q2006 puts its main config elsewhere!
    Dim sAppData As String
    Dim sIniPath As String
    sAppData = GetSpecialFolder(CSIDL_COMMON_APPDATA)
    If Dir(sAppData & "\Intuit\Quicken\Config", vbDirectory) <> "" Then
        sIniPath = sAppData & "\Intuit\Quicken\Config\"
    Else
        sIniPath = ""   ' system will use default search path for the ini file
    End If
    sIniPath = sIniPath & "QUICKEN.INI"
    
    sQuickenProg = ReadIniString(sIniPath, "Quicken", "ExePath")
    If sQuickenProg <> "" Then
        sQuickenVer = ReadIniString(sIniPath, "Quicken", "Version")
        oFile.WriteLine "Quicken found at " & sQuickenProg
        oFile.WriteLine "Version: " & sQuickenVer
        sTmp = ReadIniString(sIniPath, "Quicken", "sCurrency")
        oFile.WriteLine "sCurrency: " & sTmp
        sTmp = ReadIniString(sIniPath, "Quicken", "sDecimal")
        oFile.WriteLine "sDecimal: " & sTmp
        sTmp = ReadIniString(sIniPath, "Quicken", "sThousand")
        oFile.WriteLine "sThousand: " & sTmp
        oFile.WriteLine ""
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoAceMoney
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       16/02/2005-09:45:06
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DoAceMoney()
    Dim sTmp As String
    Dim iTmp As Integer
    Dim lTmp As Long
    Dim sPath As String
    Dim sVerFile As String
    Dim dllInfo As VS_FIXEDFILEINFO
    
    sPath = GetRegString(HKEY_CLASSES_ROOT, "AceMoney.Document\shell\open\command", "")
    If sPath <> "" Then
        iTmp = InStr(LCase$(sPath), "acemoney.exe")
        If iTmp > 0 Then
            sPath = Left$(sPath, iTmp + Len("acemoney.exe"))
            oFile.WriteLine "AceMoney found at " & sPath
            If Dir(sPath) <> "" Then
                If GetModuleInfo(sPath, dllInfo) Then
                    sVerFile = GetVersionString(dllInfo.dwFileVersionMS, dllInfo.dwFileVersionLS)
                    oFile.WriteLine "Version: " & sVerFile
                Else
                    oFile.WriteLine "No version information available"
                End If
            Else
                oFile.WriteLine sPath & " not found!!!"
            End If
            sTmp = GetRegString(HKEY_CURRENT_USER, "Software\MechCAD\AceMoney\Options", "Language")
            oFile.WriteLine "Language: " & sTmp
            lTmp = GetRegDword(HKEY_CURRENT_USER, "Software\MechCAD\AceMoney\Options", "AmountDecimalSeparator")
            oFile.WriteLine "Decimal: &H" & Hex(lTmp) & "=" & Chr$(lTmp)
            lTmp = GetRegDword(HKEY_CURRENT_USER, "Software\MechCAD\AceMoney\Options", "DateFormat")
            oFile.WriteLine "Date Format: &H" & Hex(lTmp)
            lTmp = GetRegDword(HKEY_CURRENT_USER, "Software\MechCAD\AceMoney\Options", "DateSeparator")
            oFile.WriteLine "Date Separator: &H" & Hex(lTmp) & "=" & Chr$(lTmp)
            oFile.WriteLine ""
        End If
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoMoneyDance
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       16/02/2005-21:28:53
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DoMoneyDance()
    Dim sReg As String
    Dim sTmp As String
    Dim iSlash As Integer
    sReg = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Moneydance"
    sTmp = GetRegString(HKEY_LOCAL_MACHINE, sReg, "UninstallString")
    If sTmp <> "" Then
        iSlash = InStrRev(sTmp, "\")
        sTmp = Left$(sTmp, iSlash - 1)
        iSlash = InStrRev(sTmp, "\")
        sTmp = Left$(sTmp, iSlash - 1)
        oFile.WriteLine "Moneydance found at " & sTmp
        oFile.WriteLine ""
    End If
End Sub
Private Function MonthName(iMon As Integer) As String
    Static aNames(11) As String
    Static bReady As Boolean
    
    If Not bReady Then
        aNames(0) = "Jan": aNames(1) = "Feb": aNames(2) = "Mar"
        aNames(3) = "Apr": aNames(4) = "May": aNames(5) = "Jun"
        aNames(6) = "Jul": aNames(7) = "Aug": aNames(8) = "Sep"
        aNames(9) = "Oct": aNames(10) = "Nov": aNames(11) = "Dec"
        bReady = True
    End If

    If iMon < 1 Or iMon > 12 Then
        MonthName = "?" & CStr(iMon)
    Else
        MonthName = aNames(iMon - 1)
    End If
End Function
Private Function FormatSupportDate(dDate As Date) As String
    Dim sTmp As String
'    sTmp = Format(dDate, "dd-mmm-yyyy")
    sTmp = Format(Day(dDate), "00")
    sTmp = sTmp & "-" & MonthName(Month(dDate))
    sTmp = sTmp & "-" & Format(Year(dDate), "0000")
    FormatSupportDate = sTmp
End Function
Private Function FormatSupportDateTime(dDate As Date) As String
    Dim sTmp As String
    sTmp = FormatSupportDate(dDate) & " " & Format(dDate, "Long Time")
    FormatSupportDateTime = sTmp
End Function

