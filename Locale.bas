Attribute VB_Name = "Locale"
Option Explicit

Private Const ModHeader As String = "$Header: /MT2OFX/Locale.bas 22    14/11/10 23:55 Colin $"

Const ToolTipOffset As Long = 20000

Private LocaleID As Long
Public ResOffset As Long
Public Declare Sub SetThreadLocale Lib "kernel32" (LCID As Long)
Public Declare Function GetThreadLocale Lib "kernel32" () As Long
Public Declare Function GetLocaleInfo Lib "unicows.dll" _
    Alias "GetLocaleInfoW" (ByVal Locale As Long, ByVal LCType As Long, _
    ByVal lpLCData As Long, ByVal cchData As Long) As Long
Public Declare Function EnumSystemLocales Lib "unicows.dll" _
    Alias "EnumSystemLocalesW" _
    (ByVal lpLocaleEnumProc As Long, _
    ByVal dwFlags As Long) As Long

'SYSGEOTYPE
Private Const GEO_NATION As Long = &H1
Private Const GEO_LATITUDE As Long = &H2
Private Const GEO_LONGITUDE As Long = &H3
Private Const GEO_ISO2 As Long = &H4
Private Const GEO_ISO3 As Long = &H5
Private Const GEO_RFC1766 = &H6
Private Const GEO_LCID As Long = &H7
Private Const GEO_FRIENDLYNAME As Long = &H8
Private Const GEO_OFFICIALNAME As Long = &H9
Private Const GEO_TIMEZONES As Long = &HA
Private Const GEO_OFFICIALLANGUAGES As Long = &HB

'SYSGEOCLASS
Private Const GEOCLASS_NATION As Long = 16 'only valid GeoClass value at present
Private Const GEOCLASS_REGION As Long = 14 'defined but not yet supported by Windows

Private Const GEOID_NOT_AVAILABLE As Long = -1

Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

' NB: Geo functions not available in Windows 95/98
Private Declare Function GetUserGeoID Lib "kernel32" _
  (ByVal geoclass As Long) As Long

Private Declare Function GetGeoInfo Lib "kernel32" _
   Alias "GetGeoInfoA" _
  (ByVal geoid As Long, _
   ByVal GeoType As Long, _
   lpGeoData As Any, _
   ByVal cchData As Long, _
   ByVal langid As Long) As Long

Private Declare Function EnumSystemGeoID Lib "kernel32" _
  (ByVal geoclass As Long, _
   ByVal ParentGeoId As Long, _
   ByVal lpGeoEnumProc As Long) As Long
   
Private Declare Function lstrlenW Lib "kernel32" _
  (ByVal lpString As Long) As Long

Public Const LOCALE_SLANGUAGE        As Long = &H2
Public Const LOCALE_SCOUNTRY         As Long = &H6
Public Const LOCALE_IDATE            As Long = &H21
Public Const LOCALE_ILDATE           As Long = &H22
Public Const LOCALE_SDECIMAL         As Long = &HE
Public Const LOCALE_STHOUSAND        As Long = &HF
Public Const LOCALE_SLIST            As Long = &HC
' 20081012 CS: values of LOCALE_USER_DEFAULT and LOCALE_SYSTEM_DEFAULT were swapped! How the hell did this
' persist so long...
Public Const LOCALE_USER_DEFAULT     As Long = &H400
Public Const LOCALE_SYSTEM_DEFAULT   As Long = &H800
Public Const LOCALE_SCURRENCY        As Long = &H14  'local symbol
Public Const LOCALE_SINTLSYMBOL      As Long = &H15  'intl symbol
Public Const LOCALE_SMONDECIMALSEP   As Long = &H16  'decimal separator
Public Const LOCALE_SMONTHOUSANDSEP  As Long = &H17  'thousand separator
Public Const LOCALE_SMONGROUPING     As Long = &H18  'grouping
Public Const LOCALE_ICURRDIGITS      As Long = &H19  '# local digits
Public Const LOCALE_IINTLCURRDIGITS  As Long = &H1A  '# intl digits
Public Const LOCALE_ICURRENCY        As Long = &H1B  'pos currency mode
Public Const LOCALE_INEGCURR         As Long = &H1C  'neg currency mode
Public Const LOCALE_SSHORTDATE       As Long = &H1F
Public Const LOCALE_SLONGDATE        As Long = &H20
Public Const LOCALE_IPOSSIGNPOSN     As Long = &H52  'pos sign position
Public Const LOCALE_INEGSIGNPOSN     As Long = &H53  'neg sign position
Public Const LOCALE_IPOSSYMPRECEDES  As Long = &H54  'mon sym precedes pos amt
Public Const LOCALE_IPOSSEPBYSPACE   As Long = &H55  'mon sym sep by space from pos amt
Public Const LOCALE_INEGSYMPRECEDES  As Long = &H56  'mon sym precedes neg amt
Public Const LOCALE_INEGSEPBYSPACE   As Long = &H57  'mon sym sep by space from neg amt
Public Const LOCALE_SENGCURRNAME     As Long = &H1007 'english name of currency
Public Const LOCALE_SNATIVECURRNAME  As Long = &H1008 'native name of currency
Public Const LOCALE_SENGCOUNTRY      As Long = &H1002 ' english name of country
Public Const LOCALE_SENGLANGUAGE     As Long = &H1001 ' english name of language
Public Const LOCALE_SISO639LANGNAME       As Long = &H59   'ISO abbreviated language name
Public Const LOCALE_SISO3166CTRYNAME      As Long = &H5A   'ISO abbreviated country name

Public Const LCID_INSTALLED         As Long = &H1  'installed locale ids
Public Const LCID_SUPPORTED         As Long = &H2  'supported locale ids
Public Const LCID_ALTERNATE_SORTS   As Long = &H4  'alternate sort locale ids

Public Const LANG_ENGLISH           As Long = &H9
Public Const LANG_DUTCH             As Long = &H13

'============ time zone stuff ===============
Private Const TIME_ZONE_ID_UNKNOWN As Long = 1
Private Const TIME_ZONE_ID_STANDARD As Long = 1
Private Const TIME_ZONE_ID_DAYLIGHT As Long = 2
Private Const TIME_ZONE_ID_INVALID As Long = &HFFFFFFFF

Private Type SYSTEMTIME
   wYear         As Integer
   wMonth        As Integer
   wDayOfWeek    As Integer
   wDay          As Integer
   wHour         As Integer
   wMinute       As Integer
   wSecond       As Integer
   wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
   Bias As Long
   StandardName(0 To 63) As Byte  'unicode (0-based)
   StandardDate As SYSTEMTIME
   StandardBias As Long
   DaylightName(0 To 63) As Byte  'unicode (0-based)
   DaylightDate As SYSTEMTIME
   DaylightBias As Long
End Type

Private Declare Function GetTimeZoneInformation Lib "kernel32" _
    (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
    
' Codepage stuff
Public Enum KnownCodePage
    CP_UNKNOWN = -1
    CP_ACP = 0
    CP_OEMCP = 1
    CP_MACCP = 2
    CP_THREAD_ACP = 3
    CP_SYMBOL = 42
    ' ARABIC
    CP_AWIN = 101 ' Bidi Windows codepage
    CP_709 = 102 ' MS-DOS Arabic Support CP 709
    CP_720 = 103 ' MS-DOS Arabic Support CP 720
    CP_A708 = 104 ' ASMO 708
    CP_A449 = 105 ' ASMO 449+
    CP_TARB = 106 ' MS Transparent Arabic
    CP_NAE = 107 ' Nafitha Enhanced Arabic Char Set
    CP_V4 = 108 ' Nafitha v 4.0
    CP_MA2 = 109 ' Mussaed Al Arabi (MA/2) CP 786
    CP_I864 = 110 ' IBM Arabic Supplement CP 864
    CP_A437 = 111 ' Ansi 437 codepage
    CP_AMAC = 112 ' Macintosh Code Page
    ' HEBREW
    CP_HWIN = 201 ' Bidi Windows codepage
    CP_862I = 202 ' IBM Hebrew Supplement CP 862
    CP_7BIT = 203 ' IBM Hebrew Supplement CP 862 Folded
    CP_ISO = 204 ' ISO Hebrew 8859-8 Character Set
    CP_H437 = 205 ' Ansi 437 codepage
    CP_HMAC = 206 ' Macintosh Code Page
    ' CODE PAGES
    CP_OEM_437 = 437
    CP_ARABICDOS = 708
    CP_DOS720 = 720
    CP_DOS737 = 737
    CP_DOS775 = 775
    CP_IBM850 = 850
    CP_IBM852 = 852
    CP_DOS861 = 861
    CP_DOS862 = 862
    CP_IBM866 = 866
    CP_DOS869 = 869
    CP_THAI = 874
    CP_EBCDIC = 875
    CP_JAPAN = 932
    CP_CHINA = 936
    CP_KOREA = 949
    CP_TAIWAN = 950
    ' UNICODE
    CP_UNICODELITTLE = 1200
    CP_UNICODEBIG = 1201
    ' CODE PAGES
    CP_EASTEUROPE = 1250
    CP_RUSSIAN = 1251
    CP_WESTEUROPE = 1252
    CP_GREEK = 1253
    CP_TURKISH = 1254
    CP_HEBREW = 1255
    CP_ARABIC = 1256
    CP_BALTIC = 1257
    CP_VIETNAMESE = 1258
    ' KOREAN
    CP_JOHAB = 1361
    ' MAC
    CP_MAC_ROMAN = 10000
    CP_MAC_JAPAN = 10001
    CP_MAC_ARABIC = 10004
    CP_MAC_GREEK = 10006
    CP_MAC_CYRILLIC = 10007
    CP_MAC_LATIN2 = 10029
    CP_MAC_TURKISH = 10081
    ' CODE PAGES
    CP_CHINESECNS = 20000
    CP_CHINESEETEN = 20002
    CP_IA5WEST = 20105
    CP_IA5GERMAN = 20106
    CP_IA5SWEDISH = 20107
    CP_IA5NORWEGIAN = 20108
    CP_ASCII = 20127
    CP_RUSSIANKOI8R = 20866
    CP_RUSSIANKOI8U = 21866
    CP_ISOLATIN1 = 28591
    CP_ISOEASTEUROPE = 28592
    CP_ISOTURKISH = 28593
    CP_ISOBALTIC = 28594
    CP_ISORUSSIAN = 28595
    CP_ISOARABIC = 28596
    CP_ISOGREEK = 28597
    CP_ISOHEBREW = 28598
    CP_ISOTURKISH2 = 28599
    CP_ISOLATIN9 = 28605
    CP_HEBREWLOG = 38598
    CP_USER = 50000
    CP_AUTOALL = 50001
    CP_JAPANNHK = 50220
    CP_JAPANESC = 50221
    CP_JAPANISO = 50222
    CP_KOREAISO = 50225
    CP_TAIWANISO = 50227
    CP_CHINAISO = 50229
    CP_AUTOJAPAN = 50932
    CP_AUTOCHINA = 50936
    CP_AUTOKOREA = 50949
    CP_AUTOTAIWAN = 50950
    CP_AUTORUSSIAN = 51251
    CP_AUTOGREEK = 51253
    CP_AUTOARABIC = 51256
    CP_JAPANEUC = 51932
    CP_CHINAEUC = 51936
    CP_KOREAEUC = 51949
    CP_TAIWANEUC = 51950
    CP_CHINAHZ = 52936
    CP_GB18030 = 54936
    ' UNICODE
    CP_UTF7 = 65000
    CP_UTF8 = 65001
End Enum

Private Const MAX_DEFAULTCHAR = 2
Private Const MAX_LEADBYTES = 12
Private Type CPINFO
        MaxCharSize As Long                    '  max length (Byte) of a char
        DefaultChar(1 To MAX_DEFAULTCHAR) As Byte   '  default character
        LeadByte(1 To MAX_LEADBYTES) As Byte        '  lead byte ranges
End Type
Private Type CPINFOEX
        MaxCharSize As Long                    '  max length (Byte) of a char
        DefaultChar(1 To MAX_DEFAULTCHAR) As Byte   '  default character
        LeadByte(1 To MAX_LEADBYTES) As Byte        '  lead byte ranges
        UnicodeDefaultChar As Integer          '
        CodePage As Long                 ' you read it right - this is a STRING
        CodePageName As String * 260            ' MAX_PATH
End Type

' Flags
Public Const MB_PRECOMPOSED = &H1
Public Const MB_COMPOSITE = &H2
Public Const MB_USEGLYPHCHARS = &H4
Public Const MB_ERR_INVALID_CHARS = &H8

Public Const WC_DEFAULTCHECK = &H100 ' check for default char
Public Const WC_COMPOSITECHECK = &H200 ' convert composite to precomposed
Public Const WC_DISCARDNS = &H10 ' discard non-spacing chars
Public Const WC_SEPCHARS = &H20 ' generate separate chars
Public Const WC_DEFAULTCHAR = &H40 ' replace with default char

Private Declare Function EnumSystemCodePages Lib "kernel32" Alias "EnumSystemCodePagesA" ( _
    ByVal lpCodePageEnumProc As Long, ByVal dwFlags As Long) As Long
Public Declare Function IsValidCodePage Lib "unicows.dll" (ByVal CodePage As Long) As Long

Private Const CP_INSTALLED = &H1
Private Const CP_SUPPORTED = &H2

Private lCodePages() As Long

'====================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, _
    ByVal Length As Long)

Private aCurrencyList() As String
Private iCurrencyCount As Integer

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const CCM_FIRST = &H2000&
Private Const CCM_SETUNICODEFORMAT = (CCM_FIRST + 5)
Private Const LVM_SETUNICODEFORMAT = CCM_SETUNICODEFORMAT

Private bResFileMode As Boolean
Private sResFile As String
Private Const gsResFileSection As String = "MT2OFX Language File"

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       SystemShortDateFormat
' Description:       returns the user's configured short date format
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       19/01/2004-21:21:45
'
' Parameters :
'               return:
'   Value   Meaning
'   0       Month -Day - Year
'   1       Day -Month - Year
'   2       Year -Month - Day
'--------------------------------------------------------------------------------
'</CSCM>
Public Const DATEFMT_MDY As Integer = 0
Public Const DATEFMT_DMY As Integer = 1
Public Const DATEFMT_YMD As Integer = 2
Public Const DATEFMT_SYSTEM As Integer = 3
Public Const DATEFMT_CUSTOM As Integer = 4

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetLocaleString
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       30/07/2004-21:53:51
'
' Parameters :       iWhich (Long)
'                    sDefault (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetLocaleString(iWhere As Long, iWhich As Long, sDefault As String) As String
    Dim iTmp As Long
    Dim sTmp As String
    sTmp = String(1024, vbNullChar)
    iTmp = GetLocaleInfo(iWhere, iWhich, StrPtr(sTmp), Len(sTmp))
    If iTmp > 0 Then
        GetLocaleString = Left$(sTmp, iTmp - 1) ' -1 because the count includes the NULL
    Else
        GetLocaleString = sDefault
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetLocaleInt
' Description:       Returns as locale value as an integer
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/3/2008-17:40:54
'
' Parameters :       iWhere (Long)
'                    iWhich (Long)
'                    iDefault (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetLocaleInt(iWhere As Long, iWhich As Long, iDefault As Long) As Long
    Dim iTmp As Long
    Dim sTmp As String
    sTmp = CStr(iDefault)
    sTmp = GetLocaleString(iWhere, iWhich, sTmp)
    If IsNumeric(sTmp) Then
        GetLocaleInt = CLng(sTmp)
    Else
        GetLocaleInt = iDefault
    End If
End Function

Public Function SystemShortDateFormat() As Integer
    SystemShortDateFormat = Val(GetLocaleString(LOCALE_USER_DEFAULT, LOCALE_IDATE, CStr(DATEFMT_DMY)))
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       SystemDecimalSeparator
' Description:       returns system decimal separator character
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       19/01/2004-22:31:28
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function SystemDecimalSeparator() As String
    SystemDecimalSeparator = GetLocaleString(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, ".")
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       SystemListSeparator
' Description:       Returns user's configured list separator char (usually comma or semicolon)
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       30/07/2004-21:51:35
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function SystemListSeparator() As String
    SystemListSeparator = GetLocaleString(LOCALE_USER_DEFAULT, LOCALE_SLIST, ",")
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetUserCurrency
' Description:       Returns ISO currency code for user's currency
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       07/10/2004-12:10:09
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetUserCurrency() As String
    GetUserCurrency = GetLocaleString(LOCALE_USER_DEFAULT, LOCALE_SINTLSYMBOL, "")
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetCurrencyList
' Description:       returns array of known currencies
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       07/10/2004-12:12:23
'
' Parameters :
' Notes      :          The list is not sorted, and may contain duplicates!
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetCurrencyList()
    ReDim aCurrencyList(1 To 500)
    iCurrencyCount = 0
    Call EnumSystemLocales(AddressOf EnumSystemLocalesProc, LCID_INSTALLED)
' sort and remove duplicates
    ReDim Preserve aCurrencyList(1 To iCurrencyCount)
'    QuickSortStringsAscending aCurrencyList, 1, iCurrencyCount
' remove duplicates

    GetCurrencyList = aCurrencyList
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       EnumSystemLocalesProc
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       07/10/2004-12:31:23
'
' Parameters :       lpLocaleString (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function EnumSystemLocalesProc(lpLocaleString As Long) As Long

  'application-defined callback function for EnumSystemLocales

    Dim Pos As Integer
    Dim dwLocaleDec As Long
    Dim dwLocaleHex As String
    Dim sCurrName As String
    Dim sCurrCode As String
  
  'pad a string to hold the format
    dwLocaleHex = Space$(32)
   
  'copy the string pointed to by the return value
    CopyMemory ByVal StrPtr(dwLocaleHex), lpLocaleString, ByVal Len(dwLocaleHex)
  
  'locate the terminating null
    Pos = InStr(dwLocaleHex, vbNullChar)
   
    If Pos Then
     'strip the null
        dwLocaleHex = Left$(dwLocaleHex, Pos - 1)
      
     'we need the last 4 chrs - this
     'is the locale ID in hex
        dwLocaleHex = (Right$(dwLocaleHex, 4))
      
     'convert the string to a long
        dwLocaleDec = CLng("&H" & dwLocaleHex)
      
     'get the language and abbreviation for that locale
        sCurrCode = GetUserLocaleInfo(dwLocaleDec, LOCALE_SINTLSYMBOL)
        sCurrName = GetUserLocaleInfo(dwLocaleDec, LOCALE_SENGCURRNAME)
    End If
   
  'add the data to the list
    iCurrencyCount = iCurrencyCount + 1
    aCurrencyList(iCurrencyCount) = sCurrCode & " (" & sCurrName & ")"
    
  'and return 1 to continue enumeration
    EnumSystemLocalesProc = 1
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetUserLocaleInfo
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       07/10/2004-12:43:33
'
' Parameters :       dwLocaleID (Long)
'                    dwLCType (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetUserLocaleInfo(ByVal dwLocaleID As Long, _
                                  ByVal dwLCType As Long) As String

   Dim sReturn As String
   Dim nSize As Long

  'call the function passing the Locale type
  'variable to retrieve the required size of
  'the string buffer needed
   nSize = GetLocaleInfo(dwLocaleID, dwLCType, StrPtr(sReturn), Len(sReturn))
    
  'if successful..
   If nSize Then
    
     'pad a buffer with spaces
      sReturn = Space$(nSize)
       
     'and call again passing the buffer
      nSize = GetLocaleInfo(dwLocaleID, dwLCType, StrPtr(sReturn), Len(sReturn))
     
     'if successful (nSize > 0)
      If nSize Then
      
        'nSize holds the size of the string
        'including the terminating null
         GetUserLocaleInfo = Left$(sReturn, nSize - 1)
      
      End If
   
   End If
    
End Function
Public Function GetProgLocale() As Long
    GetProgLocale = LocaleID
End Function

Public Sub SetProgLocale(ByVal l As Long)
    Select Case PRIMARYLANGID(l)
    Case LANG_DUTCH
        ResOffset = 2000
    Case Else
        ResOffset = 0
        l = &H809   ' default to English
    End Select
    LocaleID = l
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       SetLanguageFile
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       19/02/2005-12:54:58
'
' Parameters :       sFile (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function SetLanguageFile(ByVal sFile As String) As Boolean
    Dim sLCID As String
    If Len(sFile) = 0 Then
        bResFileMode = False
        SetLanguageFile = True
    Else
        If PathIsRelative(sFile) Then
            sFile = Cfg.ResourcePath & "\" & sFile
        End If
        If Dir(sFile) <> "" Then
            sResFile = sFile
            bResFileMode = True
            sLCID = ReadIniString(sFile, gsResFileSection, "LCID")
            If sLCID <> "" Then
                If IsNumeric(sLCID) Then
                    LocaleID = CLng(sLCID)
                End If
            End If
            SetLanguageFile = True
        Else
            SetLanguageFile = False
        End If
    End If
End Function

Public Function LoadResStringL(ResID As Long) As String
    On Error Resume Next
    Dim sTmp As String
    If bResFileMode Then
        sTmp = ReadIniString(sResFile, gsResFileSection, CStr(ResID))
    Else
        sTmp = LoadResString(ResID + ResOffset)
    End If
    sTmp = Replace(sTmp, "\n", vbCrLf)
    LoadResStringL = sTmp
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       LoadResStringLEx
' Description:       Returns a localised string with parameter substitution
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       04/01/2004-22:29:49
'
' Parameters :       ResID (Long)
'                    Args() (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function LoadResStringLEx(ResID As Long, ParamArray Args() As Variant) As String
    Dim sTmp As String
    Dim i As Long
    
    sTmp = LoadResStringL(ResID)
    For i = 0 To UBound(Args)
        sTmp = Replace(sTmp, "%" & CStr(i + 1), CStr(Args(i)))
    Next
    LoadResStringLEx = sTmp
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       LocaliseForm
' Description:       Localise controls on a form from resource strings
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       07/01/2004-22:56:48
'
' Parameters :       f (Form)
'                    iCaption (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function LocaliseForm(f As Form, Optional iCaption As Long = -1) As Boolean
    Dim c As Control
    Dim sTmp As String
    Dim lv As ListView
    Dim vStr As Variant
    Dim i As Integer
    Dim fnt As New StdFont
    Dim bCanProcess As Boolean
    
    On Error GoTo nextctl
    
    fnt.Name = "Segoe UI"
    fnt.size = 8.25
    
    If iCaption < 0 Then
        iCaption = f.HelpContextID
    End If
    If iCaption > 0 Then
        sTmp = LoadResStringL(iCaption)
        If Len(sTmp) > 0 Then
            f.Caption = sTmp
        End If
    End If
    For Each c In f.Controls
'        Debug.Print TypeName(c)
        Select Case TypeName(c)
        Case "Label", "TextBox", "CommandButton", "CheckBox", "ComboBox", "Frame", "OptionButton"
            If c.WhatsThisHelpID > 0 Then
                sTmp = LoadResStringL(c.WhatsThisHelpID)
                If Len(sTmp) > 0 Then
                    c.Caption = sTmp
                End If
            End If
            bCanProcess = True
        Case "ListView"
            If c.WhatsThisHelpID > 0 Then
                sTmp = LoadResStringL(c.WhatsThisHelpID)
                If Len(sTmp) > 0 Then
                    vStr = Split(sTmp, "|")
                    Set lv = c
                    For i = 1 To lv.ColumnHeaders.Count
                        lv.ColumnHeaders(i).Text = vStr(i - 1)
                    Next
                End If
            End If
            bCanProcess = True
        Case Else
            bCanProcess = False
        End Select
' now try for a tool tip
        If bCanProcess Then
'            If c.Font.Name = "MS Sans Serif" Then
'                Set c.Font = fnt
'            End If
            sTmp = LoadResStringL(c.WhatsThisHelpID + ToolTipOffset)
            If Len(sTmp) > 0 Then
                c.ToolTipText = sTmp
            End If
        End If
nextctl:
    Err.Clear
    Next
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       SUBLANGID
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       04/03/2005-22:51:10
'
' Parameters :       LCID (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function SUBLANGID(LCID As Long) As Integer
    SUBLANGID = ((LCID And &H7C) \ &H400) And &H3F
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       PRIMARYLANGID
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       04/03/2005-22:51:51
'
' Parameters :       LCID (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function PRIMARYLANGID(LCID As Long) As Integer
    PRIMARYLANGID = LCID And &H3FF
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetCurrentTimeBias
' Description:       Returns current local timezone bias in minutes
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       4/24/2006-22:25:02
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetCurrentTimeBias() As Long

   Dim tzi As TIME_ZONE_INFORMATION
   Dim dwBias As Long
'   Dim tmp As String

   Select Case GetTimeZoneInformation(tzi)
   Case TIME_ZONE_ID_DAYLIGHT
      dwBias = tzi.Bias + tzi.DaylightBias
   Case Else
      dwBias = tzi.Bias + tzi.StandardBias
   End Select

'   tmp = CStr(dwBias \ 60) & " hours, " & CStr(dwBias Mod 60) & " minutes"

   GetCurrentTimeBias = dwBias
   
End Function

Public Function InternalToLocalTime(dDate As Date) As Date
    Dim dTmp As Date
    dTmp = dDate
' if there is a non-zero time part, this time is in GMT and needs to be corrected to the local timezone
' in case this causes the date to change to the next/previous day!
    If (Hour(dTmp) + Minute(dTmp) + Second(dTmp)) > 0 Then
        dTmp = DateAdd("n", -GetCurrentTimeBias(), dTmp)
    End If
    InternalToLocalTime = dTmp
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetUserCountry
' Description:       Returns the country of the user's location as two-letter ISO code
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       1/10/2007-13:49:46
'
' Parameters :       none
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetUserCountry() As String
    Dim lGeoId As Long
    Dim lProgLcid As Long
    Dim sTmp As String
    If IsXpOrMeOrLater() Then
        lGeoId = GetUserGeoID(GEOCLASS_NATION)
        sTmp = String$(10, vbNullChar)
        GetGeoInfo lGeoId, GEO_ISO2, ByVal sTmp, Len(sTmp), 0
        GetUserCountry = LCase$(Left$(sTmp, InStr(sTmp, vbNullChar) - 1))
    Else
' 95/98/2000: get country from user locale
        lProgLcid = GetProgLocale()
        sTmp = LCase$(GetUserLocaleInfo(lProgLcid, LOCALE_SISO3166CTRYNAME))
        GetUserCountry = sTmp
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetUserLanguage
' Description:       [type_description_here]
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       22/04/2010-23:37:40
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetUserLanguage() As String
    Dim sTmp As String
    sTmp = GetLocaleString(GetProgLocale(), LOCALE_SISO639LANGNAME, "en")
    GetUserLanguage = sTmp
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetUserDateSequence
' Description:       Returns a string representing user's default date sequence (MDY, DMY, YMD)
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/3/2008-17:40:11
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetUserDateSequence() As String
    Dim iDateSeq As Long
    Dim sTmp As String
    iDateSeq = GetLocaleInt(LOCALE_USER_DEFAULT, LOCALE_ILDATE, 0) ' defaults to MDY
    Select Case iDateSeq
    Case 0
        sTmp = "MDY"
    Case 1
        sTmp = "DMY"
    Case 2
        sTmp = "YMD"
    Case Else
        sTmp = "MDY"
    End Select
    GetUserDateSequence = sTmp
End Function

Private Function IsXpOrMeOrLater() As Boolean
    Dim osvi As OSVERSIONINFOEX
    Dim osvi_old As OSVERSIONINFO   ' only used to get the size

' default to safe option! returning False will use user locale instead of geo functions
    IsXpOrMeOrLater = False
    osvi.dwOSVersionInfoSize = Len(osvi_old)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If

    If osvi.dwMajorVersion >= 6 _
    Or (osvi.dwMajorVersion = 5 And osvi.dwMinorVersion >= 1) _
    Or (osvi.dwMajorVersion = 4 And osvi.dwMinorVersion >= 90) Then
        IsXpOrMeOrLater = True
    End If
End Function

Public Function IsXpOrLater() As Boolean
    Dim osvi As OSVERSIONINFOEX
    Dim osvi_old As OSVERSIONINFO   ' only used to get the size

' default to safe option! returning False will use user locale instead of geo functions
    IsXpOrLater = False
    osvi.dwOSVersionInfoSize = Len(osvi_old)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If

    If osvi.dwMajorVersion >= 6 _
    Or (osvi.dwMajorVersion = 5 And osvi.dwMinorVersion >= 1) Then
        IsXpOrLater = True
    End If
End Function

Private Function CodePageEnumProc(CP_Pointer As Long) As Long
    Dim buffer As String
    buffer = Space$(255)
    Call CopyMemory(ByVal buffer, CP_Pointer, Len(buffer))
    buffer = Left$(buffer, InStr(buffer, vbNullChar) - 1)
    lCodePages(UBound(lCodePages)) = CLng(buffer)
    ReDim Preserve lCodePages(0 To UBound(lCodePages) + 1)
    CodePageEnumProc = 1&
End Function

Public Function GetCodePageList() As Long()
    Dim i As Long, j As Long
    Dim Flag As Long
    ReDim lCodePages(0 To 0)
    Dim asTmp() As String
    
    asTmp = EnumRegKeys(HKEY_CLASSES_ROOT, "Mime\Database\Codepage")
    
    For i = LBound(asTmp) To UBound(asTmp)
        If Len(asTmp(i)) > 0 Then
            If IsNumeric(asTmp(i)) Then
                lCodePages(UBound(lCodePages)) = CLng(asTmp(i))
                ReDim Preserve lCodePages(0 To UBound(lCodePages) + 1)
            End If
        End If
    Next
    
'    Flag = CP_INSTALLED
'    Call EnumSystemCodePages(AddressOf CodePageEnumProc, Flag)

    BubbleSort lCodePages
    GetCodePageList = lCodePages
End Function

Public Function GetCodePageName(iCP As Long) As String
    Dim sTmp As String
    Dim cpi As CPINFOEX
    Dim iTmp As Long
    Dim iRes As Long
    cpi.CodePageName = ""
'    If GetCPInfoEx(iCP, 0, cpi) = 0 Then
'        sTmp = ""
'    Else
'        iTmp = InStr(cpi.CodePageName, vbNullChar)
'        sTmp = Left$(cpi.CodePageName, iTmp - 1)
'    End If

    sTmp = GetRegStringRes(HKEY_CLASSES_ROOT, "Mime\Database\Codepage\" & CStr(iCP), "Description")
    GetCodePageName = sTmp
End Function


'Public Function ListviewSetUnicode(lv As ListView)
'    Call SendMessage(lv.hWnd, LVM_SETUNICODEFORMAT, ByVal True, 0)
'End Function


