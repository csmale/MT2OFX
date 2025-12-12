Attribute VB_Name = "Utils"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : Utils
'    Project    : MT2OFX
'
'    Description: Utility Functions
'
'    Modified   :
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/Utils.bas 37    9/01/11 11:57 Colin $"
Rem $History: Utils.bas $
' 
' *****************  Version 37  *****************
' User: Colin        Date: 9/01/11    Time: 11:57
' Updated in $/MT2OFX
'
' *****************  Version 36  *****************
' User: Colin        Date: 15/11/10   Time: 0:05
' Updated in $/MT2OFX
'
' *****************  Version 35  *****************
' User: Colin        Date: 24/11/09   Time: 22:05
' Updated in $/MT2OFX
' for 3.6 beta
'
' *****************  Version 34  *****************
' User: Colin        Date: 6/10/09    Time: 0:27
' Updated in $/MT2OFX
' a few new functions
'
' *****************  Version 33  *****************
' User: Colin        Date: 30/08/09   Time: 13:33
' Updated in $/MT2OFX
' bug with empty string while converting number
'
' *****************  Version 32  *****************
' User: Colin        Date: 15/06/09   Time: 19:25
' Updated in $/MT2OFX
' For transfer to new laptop
'
' *****************  Version 32  *****************
' User: Colin        Date: 17/01/09   Time: 23:17
' Updated in $/MT2OFX
'
' *****************  Version 30  *****************
' User: Colin        Date: 19/04/08   Time: 23:08
' Updated in $/MT2OFX
'
' *****************  Version 29  *****************
' User: Colin        Date: 7/12/06    Time: 15:07
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 28  *****************
' User: Colin        Date: 25/04/06   Time: 21:45
' Updated in $/MT2OFX
'
' *****************  Version 26  *****************
' User: Colin        Date: 1/03/06    Time: 23:12
' Updated in $/MT2OFX
'
' *****************  Version 24  *****************
' User: Colin        Date: 2/11/05    Time: 23:03
' Updated in $/MT2OFX
' V3.4 beta 1
'
' *****************  Version 23  *****************
' User: Colin        Date: 11/06/05   Time: 19:33
' Updated in $/MT2OFX
'
' *****************  Version 22  *****************
' User: Colin        Date: 8/05/05    Time: 12:43
' Updated in $/MT2OFX
' V3.3.8
'
' *****************  Version 21  *****************
' User: Colin        Date: 23/03/05   Time: 22:13
' Updated in $/MT2OFX
' Leaving for Ireland!
'
' *****************  Version 20  *****************
' User: Colin        Date: 18/03/05   Time: 21:57
' Updated in $/MT2OFX
'
' *****************  Version 19  *****************
' User: Colin        Date: 6/03/05    Time: 0:35
' Updated in $/MT2OFX
'</CSCC>

Public Declare Function HtmlHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" _
   (ByVal hwndCaller As Long, _
   ByVal pszFile As String, _
   ByVal uCommand As Long, _
   dwData As Any) As Long
Public Const HH_DISPLAY_TOC = 1
Public Const HH_DISPLAY_TOPIC = 0
Public Const HH_HELP_CONTEXT = &HF
Public Const HH_GET_LAST_ERROR = &H14

Public Type HH_LAST_ERROR_STRUCT
    cbStruct As Integer
    hr As Long
    pDescription As String    ' pointer to string?
End Type

Private Declare Function xPathIsRelative Lib "shlwapi.dll" Alias "PathIsRelativeA" _
    (ByVal sPath As String) As Long

' declarations for message box sounds
Declare Function MessageBeep Lib "User32" (ByVal wType As Long) As Long
Private Declare Function MessageBox Lib "unicows.dll" Alias "MessageBoxW" ( _
    ByVal hWnd As Long, _
    ByVal lpText As Long, _
    ByVal lpCaption As Long, _
    ByVal wType As Long) As Long

Public Const MB_OK = &H0&
Public Const MB_ICONASTERISK = &H40&
Public Const MB_ICONEXCLAMATION = &H30&
Public Const MB_ICONQUESTION = &H20&
Public Const MB_ICONHAND = &H10&
Public Const MB_TOPMOST = &H40000
Public Const WM_USER = &H400

' combobox declarations
Public Const CB_ERR = -1
Public Const CB_FINDSTRING = &H14C

Public Const CBEM_SETEXSTYLE = (WM_USER + 8)
Public Const CBEM_GETEXSTYLE = (WM_USER + 9)
' Combo box extended styles:
Public Const CBES_EX_NOEDITIMAGE = &H1& ' no image to left of edit portion
Public Const CBES_EX_NOEDITIMAGEINDENT = &H2& ' edit box and dropdown box will not display images
Public Const CBES_EX_PATHWORDBREAKPROC = &H4& ' NT only. Edit box uses \ . and / as word delimiters
'#if (_WIN32_IE >= 0x0400)
Public Const CBES_EX_NOSIZELIMIT = &H8& ' Allow combo box ex vertical size < combo, clipped.
Public Const CBES_EX_CASESENSITIVE = &H10& ' case sensitive search

Private Declare Function SendMessage Lib "User32" _
        Alias "SendMessageA" _
        (ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, _
    ByVal Length As Long)

Public Enum OutfileEnum
    OUTFILE_AUTO = 0
    OUTFILE_ASK = 1
    OUTFILE_TEMP = 2
End Enum
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Public Declare Function LoadLibrary _
    Lib "unicows.dll" _
    Alias "LoadLibraryW" ( _
        ByVal lpLibFileName As Long) As Long
Public Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
Public Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long

Public Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, ByVal lplpBuffer As Long, puLen As Long) As Long
Public Declare Function VerLanguageName Lib "version.dll" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Public Type VS_FIXEDFILEINFO
        dwSignature As Long
        dwStrucVersion As Long         '  e.g. 0x00000042 = "0.42"
        dwFileVersionMS As Long        '  e.g. 0x00030075 = "3.75"
        dwFileVersionLS As Long        '  e.g. 0x00000031 = "0.31"
        dwProductVersionMS As Long     '  e.g. 0x00030010 = "3.10"
        dwProductVersionLS As Long     '  e.g. 0x00000031 = "0.31"
        dwFileFlagsMask As Long        '  = 0x3F for version "0.42"
        dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
        dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
        dwFileType As Long             '  e.g. VFT_DRIVER
        dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
        dwFileDateMS As Long           '  e.g. 0
        dwFileDateLS As Long           '  e.g. 0
End Type
Public Type FILEINFO
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OriginalFileName As String
    ProductName As String
    ProductVersion As String
    Language As String
End Type
Public Enum VersionReturnValue
    eOK = 1
    eNoVersion = 2
End Enum

Public Declare Function GetModuleFileName _
    Lib "unicows.dll" _
    Alias "GetModuleFileNameW" ( _
        ByVal hModule As Long, _
        ByVal lpFileName As Long, _
        ByVal nSize As Long) As Long
        
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Private Declare Function LoadString Lib "unicows.dll" Alias "LoadStringW" (ByVal hInstance As Long, ByVal wID As Long, _
'    ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Private Declare Function LoadString _
    Lib "unicows.dll" _
    Alias "LoadStringW" ( _
        ByVal hInstance As Long, _
        ByVal wID As Long, _
        ByVal lpBuffer As Long, _
        ByVal nBufferMax As Long) As Long
' NB: ExpandEnvironmentStrings requires win2k+
Private Declare Function ExpandEnvironmentStrings _
    Lib "kernel32.dll" _
    Alias "ExpandEnvironmentStringsW" ( _
        ByVal psIn As Long, _
        ByVal psOut As Long, _
        ByVal lBufLen As Long) As Long
        
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

Public Function PathIsRelative(sPath As String) As Boolean
    PathIsRelative = (xPathIsRelative(sPath) <> 0)
End Function

Public Function CompressSpaces(sIn As String) As String
    Dim sTmp As String
    sTmp = sIn
    While InStr(sTmp, "  ") > 0
        sTmp = Replace(sTmp, "  ", " ")
    Wend
    CompressSpaces = sTmp
End Function
Public Function ChangeExtension(sIn As String, Ext As String)
    Dim iSlash As Long
    iSlash = InStrRev(sIn, "\")
    Dim sLastPart As String
    If iSlash = 0 Then
        sLastPart = sIn
    Else
        sLastPart = Mid$(sIn, iSlash + 1)
    End If
    Dim iDot As Long
    iDot = InStrRev(sLastPart, ".")
    If iDot = 0 Then
        If Ext = "" Then
            ChangeExtension = sIn   ' no extension and none required
        Else
            ChangeExtension = sIn & "." & Ext
        End If
    Else
        If Ext = "" Then
            ChangeExtension = Left$(sIn, iSlash) _
                & Left$(sLastPart, iDot - 1)
        Else
            ChangeExtension = Left$(sIn, iSlash) _
                & Left$(sLastPart, iDot) _
                & Ext
        End If
    End If
End Function

Public Function GetString(iMsg As Long, ParamArray iPars()) As String
    Dim sTmp As String
    sTmp = LoadResStringL(iMsg)
    Dim i As Integer
    For i = 1 To UBound(iPars) + 1
        sTmp = Replace(sTmp, "%" & CStr(i), CStr(iPars(i - 1)))
    Next
    GetString = sTmp
End Function

' 20080221 CS: required case conversion is now a parameter
Public Function PayeeSetCase(sPayee As String, iCase As Integer) As String
    Dim sTmp As String
' 20080220 CS: now uses LCase() and UCase() as these are Unicode-aware. StrConv screws up on things like
' cyrillic and greek chars.
    Select Case iCase
    Case 0
        sTmp = sPayee
    Case vbLowerCase
        sTmp = LCase$(sPayee)
    Case vbUpperCase
        sTmp = UCase$(sPayee)
    Case Else
        sTmp = StrConv(sPayee, iCase)
    End Select
    PayeeSetCase = sTmp
End Function


'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetExtension
' Description:       Extract extension from path name
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       25/11/2003-22:09:17
'
' Parameters :       sFile (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetExtension(sFile As String) As String
    Dim iTmp As Integer
    Dim sTmp As String
    iTmp = InStrRev(sFile, ".")
    If iTmp > 0 Then
        sTmp = Mid$(sFile, iTmp + 1)
    Else
        sTmp = ""
    End If
    GetExtension = sTmp
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetFilename
' Description:       Extract the filename from a full path
' Created by :       Project Administrator
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       10/5/2009-16:49:03
'
' Parameters :       sFile (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetFilename(sFile As String) As String
    Dim iTmp As Integer
    iTmp = InStrRev(sFile, "\")
    If iTmp = 0 Then
        GetFilename = sFile
    Else
        GetFilename = Mid$(sFile, iTmp + 1)
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       IsDirectory
' Description:       Returns true if the path given points to a directory (not a file)
' Created by :       Project Administrator
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       10/5/2009-16:59:01
'
' Parameters :       sPath (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function IsDirectory(sPath) As Boolean
    IsDirectory = (PathIsDirectory(sPath) <> 0)
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       AddTextToCombo
' Description:       Insert new value into combobox if it's not already there
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       17/07/2004-21:57:33
'
' Parameters :       cb (MSFORMS.ComboBox)
'                    sVal (String)
' 20110109 CS: wasn't working with MSForms 2.0 combo box, which doesn't have .hWnd or .Sorted. Had to
'   do it by hand...
'--------------------------------------------------------------------------------
'</CSCM>
Sub AddTextToCombo(cb As MSForms.ComboBox, sVal As String)
    Dim iIndex As Long
    Dim i As Long
    On Error Resume Next
    If sVal = "" Then
        Exit Sub
    End If
    iIndex = -1
    For i = 0 To cb.ListCount - 1
        If sVal = cb.List(i) Then
            iIndex = -2
            Exit For
        ElseIf sVal < cb.List(i) Then
            iIndex = i
            Exit For
        End If
    Next
    If iIndex = -1 Then
        cb.AddItem sVal
    ElseIf iIndex >= 0 Then
        cb.AddItem sVal, iIndex
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       FindInCombo
' Description:       find a string in a combobox/dropdown
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       17/12/2003-21:09:09
'
' Parameters :       cb (ComboBox)
'                    sString (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function FindInCombo(cb As ComboBox, sString As String) As Long
    Dim i As Long
    Dim l As Long
    l = Len(sString)
    For i = 0 To cb.ListCount - 1
        If Left$(cb.List(i), l) = sString Then
            FindInCombo = i
            Exit Function
        End If
    Next
    FindInCombo = -1
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       MakeComboCaseSensitive
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       17/07/2004-22:41:37
'
' Parameters :       cb (ComboBox)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub MakeComboCaseSensitive(cb As ComboBox)
    Dim lRet As Long
    lRet = SendMessage(cb.hWnd, CBEM_SETEXSTYLE, CBES_EX_CASESENSITIVE, CBES_EX_CASESENSITIVE)
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       EditScript
' Description:       Edit a script file, either by notepad or the
'                    the configured editor for .VBS file
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-21:14:04
'
' Parameters :       sScript (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub EditScript(sScript As String)
    On Error GoTo shellerror
    Dim sTmp As String
    Dim lPID As Long
    
    sTmp = FindScript(sScript)
' 20060125CS: FindScript checks whether the file exists
    If sTmp = "" Then
        MyMsgBox "Script file not found: " & sScript
        Exit Sub
    End If
    lPID = Shell("""" & Cfg.ScriptEditor & """ """ & sTmp & """", vbNormalFocus)
    Exit Sub
shellerror:
    MyMsgBox LoadResStringL(131) & Err.Description, vbOKOnly + vbCritical
End Sub

Public Sub DisplayOpenWith(strFile As String)

'***PURPOSE: DISPLAY OPEN WITH DIALOG:
'   PASS IT A FILE NAME
'   e.g., DisplayOpenWith "C:\FileWithNoDefaultApplication.bvq"
'**************************************

    On Error Resume Next
    Shell "rundll32.exe shell32.dll, OpenAs_RunDLL " & strFile
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       FormatAmount
' Description:       Format an amount independant of the user's
'                    regional settings, with the minus sign on the left
'                    and the given char as decimal separator
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       20/01/2004-11:49:52
'
' Parameters :       Amount (Double)
'                    DecSeparator (String)
'--------------------------------------------------------------------------------
'</CSCM>
' 20050608 CS: Now handles up to six decimal places (for e-gold), removing trailing zeros
Function FormatAmount(Amount As Double, DecSeparator As String) As String
    Dim sTmp As String
' 20050706 CS: fixed number formatting in non-default locales
' NB: Format() uses the Default System Locale, not the Default User Locale!
    sTmp = Format(Amount, "0.000000")
' correct the separator if needed
    Dim sOtherSep As String
    If InStr(sTmp, DecSeparator) = 0 Then
        sOtherSep = IIf(DecSeparator = ".", ",", ".")
        sTmp = Replace$(sTmp, sOtherSep, DecSeparator)
    End If
'    Dim sSysSep As String
'    sSysSep = SystemDecimalSeparator()
'    If DecSeparator <> sSysSep Then
'        sTmp = Replace(sTmp, sSysSep, DecSeparator)
'    End If
' 20060317 CS: ensure at least two decimal places remain to avoid problems with Money 2006 in locales which
' use the comma as decimal separator
    Dim i As Long
    For i = Len(sTmp) To InStr(sTmp, DecSeparator) + 3 Step -1
        If Mid$(sTmp, i, 1) <> "0" Then
            Exit For
        End If
    Next
    FormatAmount = Left$(sTmp, i)
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       NumberFromString
' Description:       Translate string to number
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       29/02/2004-23:19:36
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function NumberFromString(sIn As String, dec As String) As Double
    Dim iPoint As Integer
    Dim dPower As Double
    Dim iSign As Single
    Dim sOther As String
    Dim s As String
    On Error GoTo baleout
' 20050508 CS: Strip out blanks here - French banks tend to leave them in as
' a thousands separator
    s = Replace(sIn, " ", "")
' 20050709 CS: Also strip non-breaking spaces
    s = Replace(s, Chr$(160), "")
' 20050912 CS: The Swiss sometimes use a quote as a thousands separator...
    s = Replace(s, "'", "")
' 20060124 CS: Strip common currency signs
    s = Replace(s, "$", "")
    s = Replace(s, "€", "")
    s = Replace(s, "£", "")
    s = Replace(s, "¥", "")
    If Len(s) = 0 Then
        NumberFromString = 0#
        Exit Function
    End If
    If dec = "." Then
        sOther = ","
    Else
        sOther = "."
    End If
' 20050707 CS: replace other char now so left() etc work properly
    s = Replace(s, sOther, "")
' 20050911 CS: Move iSign calculation to here and clear out "-" so "-.10" works properly
' 8 Oct 2004: cannot use Sgn() as this can return 0, which will cause any fraction
' to be lost.
    iSign = IIf(InStr(s, "-") > 0, -1, 1)
' 20090716 CS: handle amounts in parentheses. note this reverses the sign, so (-10) is positive.
    iSign = iSign * IIf(InStr(s, "(") > 0, -1, 1)
    s = Replace$(s, "-", "")
    s = Replace$(s, "+", "")
    s = Replace$(s, "(", "")
    s = Replace$(s, ")", "")
    iPoint = InStr(s, dec)
    If iPoint = 0 Then
' 20090821 CS: watch for empty string
        If Len(s) = 0 Then
            NumberFromString = 0#
        Else
            NumberFromString = CDbl(s) * iSign
        End If
    Else
        If iPoint > 1 Then
            NumberFromString = CDbl(Left(s, iPoint - 1))
        Else
            NumberFromString = 0#
        End If
        NumberFromString = Abs(NumberFromString)
        iPoint = iPoint + 1
        dPower = 0.1
        While iPoint <= Len(s)
            NumberFromString = NumberFromString + (Val(Mid$(s, iPoint, 1)) * dPower)
            iPoint = iPoint + 1
            dPower = dPower / 10#
        Wend
        NumberFromString = NumberFromString * iSign
    End If
    Exit Function
baleout:
    ShowError "NumberFromString", "Problem converting '" & sIn & "'"
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       MyMsgBox
' Description:       Wrapper for MsgBox including system sounds
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       13/03/2004-22:07:09
'
' Parameters :       sString (String)
'                    iFlags (Long)
'                    sTitle (String)
'--------------------------------------------------------------------------------
'</CSCM>
' 20060815 CS: temporarily remove topmost status from progress indicator to allow the message box to
' overlay it
Public Function MyMsgBox(Prompt As String, _
    Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
    Optional Title As String = "MT2OFX", _
    Optional HelpFile As String = "MT2OFX.CHM", _
    Optional Context As Long = 0) As VbMsgBoxResult
    
    Dim bProgressOnTop As Boolean
    Dim iLogEventType As Integer
    
    bProgressOnTop = SetProgressTopmost(False)
    
    Dim iBeep As Long
    If Cfg.EnableSounds Then
        If (Buttons And 112) = vbCritical Then
            iBeep = MB_ICONHAND
        ElseIf (Buttons And 112) = vbInformation Then
            iBeep = MB_ICONASTERISK
        ElseIf (Buttons And 112) = vbExclamation Then
            iBeep = MB_ICONEXCLAMATION
        ElseIf (Buttons And 112) = vbQuestion Then
            iBeep = MB_ICONQUESTION
        Else
            iBeep = MB_OK
        End If
        MessageBeep iBeep
    End If
'    MyMsgBox = MsgBox(Prompt, (Buttons Or MB_TOPMOST), Title, HelpFile, Context)
    If CmdParams.Quiet Then
        If (Buttons And 112) = vbCritical Then
            iLogEventType = EVENTLOG_ERROR_TYPE
        ElseIf (Buttons And 112) = vbExclamation Then
            iLogEventType = EVENTLOG_WARNING_TYPE
        Else
            iLogEventType = EVENTLOG_INFORMATION_TYPE
        End If
        Call RecordEvent(iLogEventType, 1, Prompt)
        MyMsgBox = vbCancel
    Else
        MyMsgBox = MessageBox(0&, StrPtr(Prompt), StrPtr(Title), (Buttons Or MB_TOPMOST))
    End If
    bProgressOnTop = SetProgressTopmost(bProgressOnTop)
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ShowHelpDoCommand
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/08/2004-17:24:13
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub ShowHelpDoCommand(p1 As Long, p2 As Long)
    Dim iRet As Long
    Dim sTmp As String
    iRet = HtmlHelp(GetDesktopWindow(), App.HelpFile, p1, ByVal p2)
    If iRet = 0 Then
        Dim e As HH_LAST_ERROR_STRUCT
        e.cbStruct = Len(e)
        iRet = HtmlHelp(GetDesktopWindow(), vbNullString, HH_GET_LAST_ERROR, e)
        sTmp = StrConv(e.pDescription, vbFromUnicode)
        MyMsgBox sTmp, vbOKOnly + vbCritical, "Unable to display help"
    End If
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ShowHelpTopic
' Description:       open help file to given topic
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       20/08/2004-17:22:16
'
' Parameters :       Topic (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Sub ShowHelpTopic(Topic As Long)
    ShowHelpDoCommand HH_HELP_CONTEXT, Topic
End Sub

Sub ShowHelpContents()
    ShowHelpDoCommand HH_DISPLAY_TOC, 0&
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetModuleVersion
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       22/10/2004-22:24:39
'
' Parameters :       sDll (String)
'--------------------------------------------------------------------------------
'</CSCM>
Function GetModuleInfo(sDLL As String, Info As VS_FIXEDFILEINFO) As Boolean
    GetModuleInfo = False
    Dim hInfoLen As Long
    Dim hHandle As Long
    hInfoLen = GetFileVersionInfoSize(sDLL, hHandle)
    If hInfoLen = 0 Then
        ShowDllError "GetModuleInfo", "GetFileInfoSize(" & sDLL & ")"
        Exit Function
    End If
    
'    MsgBox "info struct length: " & CStr(hInfoLen)
    Dim iaVer() As Byte
    ReDim iaVer(1 To hInfoLen)
    
    If GetFileVersionInfo(sDLL, hHandle, hInfoLen, iaVer(1)) = 0 Then
        ShowDllError "GetModuleInfo", "GetFileInfoVersion(" & sDLL & ")"
        Exit Function
    End If
    Dim pValue As Long
    Dim pLen As Long
    Dim iaFixedInfo As VS_FIXEDFILEINFO
    pValue = VarPtr(iaFixedInfo)
    If VerQueryValue(iaVer(1), "\", VarPtr(pValue), pLen) = 0 Then
        ShowDllError "GetModuleInfo", "VerQueryValue(" & sDLL & ")"
        Exit Function
    End If
    CopyMemory Info, ByVal pValue, pLen
    GetModuleInfo = True
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetFileVersionInformation
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       11/03/2005-21:53:20
'
' Parameters :       pstrFieName (String)
'                    tFileInfo (FILEINFO)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetFileVersionInformation(ByRef pstrFileName As String, ByRef tFileInfo As FILEINFO) As VersionReturnValue

    Dim lBufferLen As Long, lDummy As Long
    Dim sBuffer() As Byte
    Dim lVerPointer As Long
    Dim lRet As Long
    Dim Lang_Charset_String As String
    Dim HexNumber As Long
    Dim lLang As Long
    Dim sLang As String
    Dim i As Integer
    Dim strTemp As String
    'Clear the Buffer tFileInfo
    tFileInfo.CompanyName = ""
    tFileInfo.FileDescription = ""
    tFileInfo.FileVersion = ""
    tFileInfo.InternalName = ""
    tFileInfo.LegalCopyright = ""
    tFileInfo.OriginalFileName = ""
    tFileInfo.ProductName = ""
    tFileInfo.ProductVersion = ""
    lBufferLen = GetFileVersionInfoSize(pstrFileName, lDummy)


    If lBufferLen < 1 Then
        GetFileVersionInformation = eNoVersion
        Exit Function
    End If

    ReDim sBuffer(lBufferLen)
    lRet = GetFileVersionInfo(pstrFileName, 0&, lBufferLen, sBuffer(0))


    If lRet = 0 Then
        GetFileVersionInformation = eNoVersion
        Exit Function
    End If

    lRet = VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", VarPtr(lVerPointer), lBufferLen)


    If lRet = 0 Then
        GetFileVersionInformation = eNoVersion
        Exit Function
    End If

    Dim bytebuffer(255) As Byte
    MoveMemory bytebuffer(0), lVerPointer, lBufferLen
    HexNumber = bytebuffer(2) + bytebuffer(3) * &H100 + bytebuffer(0) * &H10000 + bytebuffer(1) * &H1000000
    Lang_Charset_String = Hex(HexNumber)
    'Pull it all apart:
    '04------= SUBLANG_ENGLISH_USA
    '--09----= LANG_ENGLISH
    ' ----04E4 = 1252 = Codepage for Windows
    '     :Multilingual

    Do While Len(Lang_Charset_String) < 8
        Lang_Charset_String = "0" & Lang_Charset_String
    Loop

    sLang = "&H" & Left$(Lang_Charset_String, 4)
    If IsNumeric(sLang) Then
        lLang = CLng(sLang)
        sLang = String$(512, vbNullChar)
        lRet = VerLanguageName(lLang, sLang, Len(sLang))
        sLang = Left$(sLang, InStr(sLang, vbNullChar) - 1)
    End If
    tFileInfo.Language = sLang
    
    Dim strVersionInfo(7) As String
    strVersionInfo(0) = "CompanyName"
    strVersionInfo(1) = "FileDescription"
    strVersionInfo(2) = "FileVersion"
    strVersionInfo(3) = "InternalName"
    strVersionInfo(4) = "LegalCopyright"
    strVersionInfo(5) = "OriginalFileName"
    strVersionInfo(6) = "ProductName"
    strVersionInfo(7) = "ProductVersion"
    Dim buffer As String

    For i = 0 To 7
        buffer = String(255, 0)
        strTemp = "\StringFileInfo\" & Lang_Charset_String _
        & "\" & strVersionInfo(i)
        lRet = VerQueryValue(sBuffer(0), strTemp, VarPtr(lVerPointer), lBufferLen)


        If lRet = 0 Then
            GetFileVersionInformation = eNoVersion
            Exit Function
        End If

        lstrcpy buffer, lVerPointer
        buffer = Mid$(buffer, 1, InStr(buffer, vbNullChar) - 1)


        Select Case i
            Case 0
            tFileInfo.CompanyName = buffer
            Case 1
            tFileInfo.FileDescription = buffer
            Case 2
            tFileInfo.FileVersion = buffer
            Case 3
            tFileInfo.InternalName = buffer
            Case 4
            tFileInfo.LegalCopyright = buffer
            Case 5
            tFileInfo.OriginalFileName = buffer
            Case 6
            tFileInfo.ProductName = buffer
            Case 7
            tFileInfo.ProductVersion = buffer
        End Select
    Next i

GetFileVersionInformation = eOK
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetModulePath
' Description:       Returns full path of given DLL which is/will be loaded
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       23/10/2004-14:17:21
'
' Parameters :       sDll (String)
'--------------------------------------------------------------------------------
'</CSCM>
Function GetModulePath(sDLL As String) As String
    Dim hModule As Long
    Dim sDLLz As String
    sDLLz = sDLL & Chr$(0)
    hModule = LoadLibrary(StrPtr(sDLL))
    If hModule = 0 Then
        ShowDllError "GetModulePath", "LoadLibrary(" & sDLL & ") failed"
        Exit Function
    End If
    Dim sPath As String
    sPath = String(1024, vbNullChar)
    Dim iLen As Long
    iLen = GetModuleFileName(hModule, StrPtr(sPath), Len(sPath))
    If iLen = 0 Then
        ShowDllError "GetModulePath", "GetModuleFileName(" & sDLL & ") failed"
    End If
    FreeLibrary hModule
    GetModulePath = Left$(sPath, iLen)
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetVersionString
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       23/10/2004-14:32:53
'
' Parameters :       dwVersionMS (Variant)
'                    dwVersionLS (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Function GetVersionString(dwVersionMS As Long, dwVersionLS As Long) As String
    GetVersionString = CStr(dwVersionMS \ &H10000) _
        & "." _
        & CStr(dwVersionMS And (CLng(&H10000) - 1))
    If dwVersionLS <> -1 Then
        GetVersionString = GetVersionString _
            & "." _
            & CStr(dwVersionLS \ &H10000) _
            & "." _
            & CStr(dwVersionLS And (CLng(&H10000) - 1))
    End If
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CheckScriptingVersion
' Description:       Check version of vbscript.dll
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       22/10/2004-16:13:02
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Function CheckScriptingVersion(sDLL As String, MinVerMS As Long) As Boolean
    Dim sTmp As String
    Dim sVerFile As String
    Dim sVerMin As String
    Dim dllInfo As VS_FIXEDFILEINFO
    CheckScriptingVersion = GetModuleInfo(sDLL, dllInfo)
    If Not CheckScriptingVersion Then
        Exit Function
    End If
    If dllInfo.dwFileVersionMS < MinVerMS Then
        sTmp = GetModulePath(sDLL)
        sVerFile = GetVersionString(dllInfo.dwFileVersionMS, dllInfo.dwFileVersionLS)
        sVerMin = GetVersionString(MinVerMS, -1)
        LogMessage True, True, LoadResStringLEx(500, sDLL, sTmp, sVerFile, sVerMin), "MT2OFX"
'        LogMessage True, True, "Your version of VBScript.dll (" _
'            & sTmp _
'            & ") is too old. The file has version " _
'            & sVerFile _
'            & "; MT2OFX requires a minimum of " _
'            & sVerMin _
'            & ".", "MT2OFX"
    Else
        CheckScriptingVersion = True
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetOutputFile
' Description:       Ask the user for an output file or derive it
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       03/11/2004-12:48:00
'
' Parameters :       sIn (String) Input file name
'--------------------------------------------------------------------------------
'</CSCM>
Function GetOutputFile(sIn As String, hWnd As Long, sType As String, iOutType As Long) As String
    Dim sOut As String
' 20041224 CS changed clipboard support - only ASK and TEMP are relevant!
    Dim sTmpIn As String
    Dim dTmp As Date
    Dim iType As OutfileEnum
    If Len(sIn) = 0 Then
        If Cfg.SaveClipboardOutput Then
            iType = OUTFILE_ASK
            dTmp = Now()
            sTmpIn = Format(dTmp, "yyyymmdd-hhmmss")    ' should be config
            sTmpIn = ChangeExtension(sTmpIn, Cfg.OutputFileType)
        Else
            sTmpIn = sIn
            iType = OUTFILE_TEMP
        End If
    Else
        sTmpIn = sIn
        iType = Cfg.PromptForOutput
    End If
    Select Case iType
    Case OUTFILE_ASK
        sOut = GetOutputFileName(sTmpIn, hWnd, sType, iOutType)
    Case OUTFILE_AUTO
        sOut = ChangeExtension(sIn, Cfg.OutputFileType)
' 3 Nov 2004 If out=in we prompt for an output file anyway
        If UCase$(sIn) = UCase$(sOut) Then
            sOut = GetOutputFileName(sIn, hWnd, sType, iType)
        End If
    Case OUTFILE_TEMP
        sOut = GetTempFile(Cfg.OutputFileType)
    Case Else
        LogMessage False, True, "Unexpected PromptForOutput=" & Cfg.PromptForOutput & " in GetOutputFile"
        Debug.Assert False
        sOut = GetOutputFileName(sIn, hWnd, sType, iType)
    End Select
    GetOutputFile = sOut
End Function

Public Function GetDefaultOutputType(sFile As String) As String
    Dim sExt As String
    sExt = UCase(GetExtension(sFile))
    Select Case sExt
    Case "QIF", "OFX", "OFC", "QFX"
        GetDefaultOutputType = sExt
    Case Else
        GetDefaultOutputType = "OFX"
    End Select
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       InExtensionList
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       18-Dec-2004-00:35:46
'
' Parameters :       sExts (String)
'                    sThisExt (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function InExtensionList(sExtList As String, sThisExt As String) As Boolean
    InExtensionList = False
    Dim sExt As String
    sExt = UCase$(sThisExt)
    Dim vExts As Variant
    vExts = Split(UCase$(sExtList), ",") ' list of MT940 extensions
    Dim i As Integer
    If TypeName(vExts) = "String()" Then
        For i = LBound(vExts) To UBound(vExts)
            If sExt = vExts(i) Then
                InExtensionList = True
                Exit Function
            End If
        Next
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       AddToExtensionList
' Description:       Adds an extension to the list if it is not already there
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       18-Dec-2004-00:38:53
'
' Parameters :       sExtList (String)
'                    sThisExt (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function AddToExtensionList(sExtList As String, sThisExt As String) As Boolean
    AddToExtensionList = False
    If Not InExtensionList(sExtList, sThisExt) Then
        sExtList = sExtList & "," & UCase$(sThisExt)
        AddToExtensionList = True
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       QuickSortStringsAscending
' Description:       Sorts a string array using QuickSort
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       3/22/2006-22:58:06
'
' Parameters :       sarray() (String)
'                    inLow (Long)
'                    inHi (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub QuickSortStringsAscending(sarray() As String, inLow As Long, inHi As Long)
  
   Dim pivot As String
   Dim tmpSwap As String
   Dim tmpLow As Long
   Dim tmpHi As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = sarray((inLow + inHi) / 2)
  
   While (tmpLow <= tmpHi)
   
      While (sarray(tmpLow) < pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot < sarray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = sarray(tmpLow)
         sarray(tmpLow) = sarray(tmpHi)
         sarray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
   
   Wend
  
   If (inLow < tmpHi) Then QuickSortStringsAscending sarray(), inLow, tmpHi
   If (tmpLow < inHi) Then QuickSortStringsAscending sarray(), tmpLow, inHi
  
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       BubbleSort
' Description:       sort an array using bubble sort
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       1/17/2009-23:12:54
'
' Parameters :       arr (Variant)
'                    numEls (Variant)
'                    descending (Boolean)
'--------------------------------------------------------------------------------
'</CSCM>
Sub BubbleSort(arr As Variant, Optional numEls As Variant, _
    Optional descending As Boolean)

    Dim Value As Variant
    Dim Index As Long
    Dim firstItem As Long
    Dim indexLimit As Long, lastSwap As Long

    ' account for optional arguments
    If IsMissing(numEls) Then numEls = UBound(arr)
    firstItem = LBound(arr)
    lastSwap = numEls

    Do
        indexLimit = lastSwap - 1
        lastSwap = 0
        For Index = firstItem To indexLimit
            Value = arr(Index)
            If (Value > arr(Index + 1)) Xor descending Then
                ' if the items are not in order, swap them
                arr(Index) = arr(Index + 1)
                arr(Index + 1) = Value
                lastSwap = Index
            End If
        Next
    Loop While lastSwap
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ExpandStrings
' Description:       Replace variable references as per Windows command shell: %A% etc.
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       1/17/2009-23:11:23
'
' Parameters :       sIn (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ExpandStrings(sIn As String) As String
    Dim sTmp As String
    Dim lBufLen As Long
    sTmp = ""
    lBufLen = 0
    lBufLen = ExpandEnvironmentStrings(StrPtr(sIn), 0&, lBufLen)
    sTmp = String$(lBufLen, Chr(0))
    lBufLen = ExpandEnvironmentStrings(StrPtr(sIn), StrPtr(sTmp), lBufLen)
    ExpandStrings = Left$(sTmp, lBufLen - 1)
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       LoadStringFromDll
' Description:       Load a string resource from a DLL (or other executable such as EXE, OCX)
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       1/17/2009-23:10:04
'
' Parameters :       sDLL (String)
'                    iResource (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function LoadStringFromDll(sDLL As String, iResource As Long) As String
    Dim hModule As Long
    Dim sTmp As String
    Dim sDLLz As String
    Dim iTmp As Long
    Dim iBufLen As Long
    LoadStringFromDll = ""

    sDLLz = sDLL & Chr$(0)
    hModule = LoadLibrary(StrPtr(sDLL))
    If hModule = 0 Then
        LoadStringFromDll = sDLL
    Else
        sTmp = String$(1024, Chr(0))
        iBufLen = Len(sTmp)
        iTmp = LoadString(hModule, iResource, StrPtr(sTmp), iBufLen)
        If iTmp > 0 Then
            LoadStringFromDll = Left$(sTmp, iTmp)
        End If
        FreeLibrary hModule
    End If
End Function

