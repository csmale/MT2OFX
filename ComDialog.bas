Attribute VB_Name = "ComDialog"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : ComDialog
'    Project    : MT2OFX
'
'    Description: Common Dialog code
'
'    Modified   : $Author: Colin $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/ComDialog.bas 11    15/11/10 0:15 Colin $"
' $History: ComDialog.bas $
' 
' *****************  Version 11  *****************
' User: Colin        Date: 15/11/10   Time: 0:15
' Updated in $/MT2OFX
'
' *****************  Version 10  *****************
' User: Colin        Date: 6/10/09    Time: 0:34
' Updated in $/MT2OFX
' added watcher support for explicit output type
'
' *****************  Version 8  *****************
' User: Colin        Date: 7/12/06    Time: 14:53
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2

'</CSCC>

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2004 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const OFN_ALLOWMULTISELECT As Long = &H200
Public Const OFN_CREATEPROMPT As Long = &H2000
Public Const OFN_ENABLEHOOK As Long = &H20
Public Const OFN_ENABLETEMPLATE As Long = &H40
Public Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
Public Const OFN_EXPLORER As Long = &H80000
Public Const OFN_EXTENSIONDIFFERENT As Long = &H400
Public Const OFN_FILEMUSTEXIST As Long = &H1000
Public Const OFN_HIDEREADONLY As Long = &H4
Public Const OFN_LONGNAMES As Long = &H200000
Public Const OFN_NOCHANGEDIR As Long = &H8
Public Const OFN_NODEREFERENCELINKS As Long = &H100000
Public Const OFN_NOLONGNAMES As Long = &H40000
Public Const OFN_NONETWORKBUTTON As Long = &H20000
Public Const OFN_NOREADONLYRETURN As Long = &H8000& 'see comments
Public Const OFN_NOTESTFILECREATE As Long = &H10000
Public Const OFN_NOVALIDATE As Long = &H100
Public Const OFN_OVERWRITEPROMPT As Long = &H2
Public Const OFN_PATHMUSTEXIST As Long = &H800
Public Const OFN_READONLY As Long = &H1
Public Const OFN_SHAREAWARE As Long = &H4000
Public Const OFN_SHAREFALLTHROUGH As Long = 2
Public Const OFN_SHAREWARN As Long = 0
Public Const OFN_SHARENOWARN As Long = 1
Public Const OFN_SHOWHELP As Long = &H10
Public Const OFS_MAXPATHNAME As Long = 260

'OFS_FILE_OPEN_FLAGS and OFS_FILE_SAVE_FLAGS below
'are mine to save long statements; they're not
'a standard Win32 type.
Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_CREATEPROMPT _
             Or OFN_NODEREFERENCELINKS

Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_OVERWRITEPROMPT _
             Or OFN_HIDEREADONLY

Public Type myOPENFILENAME
  nStructSize       As Long
  hwndOwner         As Long
  hInstance         As Long
  Filter            As String
  sCustomFilter     As String
  nMaxCustFilter    As Long
  FilterIndex       As Long
  FileName          As String
  nMaxFile          As Long
  FileTitle         As String
  nMaxTitle         As Long
  InitialDir        As String
  DialogTitle       As String
  Flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  DefFileExt        As String
  nCustData         As Long
  fnHook            As Long
  sTemplateName     As String
End Type

Public Type myOPENFILENAMEW
  nStructSize       As Long
  hwndOwner         As Long
  hInstance         As Long
  Filter            As Long
  sCustomFilter     As Long
  nMaxCustFilter    As Long
  FilterIndex       As Long
  FileName          As Long
  nMaxFile          As Long
  FileTitle         As Long
  nMaxTitle         As Long
  InitialDir        As Long
  DialogTitle       As Long
  Flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  DefFileExt        As Long
  nCustData         As Long
  fnHook            As Long
  sTemplateName     As Long
End Type

Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Private Type OFNOTIFY
    hdr As NMHDR
    lpOFN As Long           ' Long pointer to OFN structure
    pszFile As String ';        // May be NULL
End Type

Public Const WM_NOTIFY As Long = &H4E

Private Const H_MAX As Long = &HFFFF + 1
Private Const CDN_FIRST = (H_MAX - 601)
Private Const CDN_LAST = (H_MAX - 699)

'// Notifications when Open or Save dialog status changes
Private Const CDN_INITDONE = (CDN_FIRST - &H0)
Private Const CDN_SELCHANGE = (CDN_FIRST - &H1)
Private Const CDN_FOLDERCHANGE = (CDN_FIRST - &H2)
Private Const CDN_SHAREVIOLATION = (CDN_FIRST - &H3)
Private Const CDN_HELP = (CDN_FIRST - &H4)
Private Const CDN_FILEOK = (CDN_FIRST - &H5)
Private Const CDN_TYPECHANGE = (CDN_FIRST - &H6)
Private Const CDN_INCLUDEITEM = (CDN_FIRST - &H7)

Public Declare Function GetOpenFileNameANSI Lib "comdlg32" _
    Alias "GetOpenFileNameA" _
   (pOpenfilename As myOPENFILENAME) As Long
   
#If UNICOWS Then
Public Declare Function GetOpenFileName Lib "UNICOWS" _
    Alias "GetOpenFileNameW" _
   (pOpenfilename As myOPENFILENAMEW) As Long
#Else
Public Declare Function GetOpenFileName Lib "comdlg32" _
    Alias "GetOpenFileNameW" _
   (pOpenfilename As myOPENFILENAMEW) As Long
#End If

Public Declare Function GetSaveFileNameANSI Lib "comdlg32" _
   Alias "GetSaveFileNameA" _
  (pOpenfilename As myOPENFILENAME) As Long

#If UNICOWS Then
Public Declare Function GetSaveFileName Lib "UNICOWS" _
   Alias "GetSaveFileNameW" _
  (pOpenfilename As myOPENFILENAMEW) As Long
#Else
Public Declare Function GetSaveFileName Lib "comdlg32" _
   Alias "GetSaveFileNameW" _
  (pOpenfilename As myOPENFILENAMEW) As Long
#End If

Public Declare Function GetShortPathName Lib "kernel32" _
    Alias "GetShortPathNameA" _
   (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, _
    ByVal Length As Long)

Private Declare Function StrLenW Lib "kernel32.dll" Alias "lstrlenW" (ByVal Ptr As Long) As Long

' Common Dialog Errors
Public Declare Function CommDlgExtendedError Lib "comdlg32" () As Long
Public Enum EDialogError
    CDERR_DIALOGFAILURE = &HFFFF

    CDERR_GENERALCODES = &H0
    CDERR_STRUCTSIZE = &H1
    CDERR_INITIALIZATION = &H2
    CDERR_NOTEMPLATE = &H3
    CDERR_NOHINSTANCE = &H4
    CDERR_LOADSTRFAILURE = &H5
    CDERR_FINDRESFAILURE = &H6
    CDERR_LOADRESFAILURE = &H7
    CDERR_LOCKRESFAILURE = &H8
    CDERR_MEMALLOCFAILURE = &H9
    CDERR_MEMLOCKFAILURE = &HA
    CDERR_NOHOOK = &HB
    CDERR_REGISTERMSGFAIL = &HC

    PDERR_PRINTERCODES = &H1000
    PDERR_SETUPFAILURE = &H1001
    PDERR_PARSEFAILURE = &H1002
    PDERR_RETDEFFAILURE = &H1003
    PDERR_LOADDRVFAILURE = &H1004
    PDERR_GETDEVMODEFAIL = &H1005
    PDERR_INITFAILURE = &H1006
    PDERR_NODEVICES = &H1007
    PDERR_NODEFAULTPRN = &H1008
    PDERR_DNDMMISMATCH = &H1009
    PDERR_CREATEICFAILURE = &H100A
    PDERR_PRINTERNOTFOUND = &H100B
    PDERR_DEFAULTDIFFERENT = &H100C

    CFERR_CHOOSEFONTCODES = &H2000
    CFERR_NOFONTS = &H2001
    CFERR_MAXLESSTHANMIN = &H2002

    FNERR_FILENAMECODES = &H3000
    FNERR_SUBCLASSFAILURE = &H3001
    FNERR_INVALIDFILENAME = &H3002
    FNERR_BUFFERTOOSMALL = &H3003

    CCERR_CHOOSECOLORCODES = &H5000
End Enum

Public Function FARPROC(ByVal pfn As Long) As Long
  
  'Dummy procedure that receives and returns
  'the return value of the AddressOf operator.
 
  'Obtain and set the address of the callback
  'This workaround is needed as you can't assign
  'AddressOf directly to a member of a user-
  'defined type, but you can assign it to another
  'long and use that (as returned here)
   FARPROC = pfn

End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OFNHookProc
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       31/01/2004-22:28:27
'
' Parameters :       hwnd (Long)
'                    uMsg (Long)
'                    wParam (Long)
'                    lParam (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function OFNHookProc(ByVal hWnd As Long, _
                            ByVal uMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long
                                   
  'On initialization, set aspects of the
  'dialog that are not obtainable through
  'manipulating the OPENFILENAME structure members.
  
   Dim hwndParent As Long
    Dim tNMH As NMHDR
    Dim tOFNOTIFY As OFNOTIFY
    Dim tOFN As myOPENFILENAME
    Dim vFilter As Variant
    Dim sTmp As String
    
    Select Case uMsg
    Case WM_NOTIFY
        CopyMemory tNMH, ByVal lParam, Len(tNMH)
        Select Case tNMH.code
        Case CDN_TYPECHANGE
            CopyMemory tOFNOTIFY, ByVal lParam, Len(tOFNOTIFY)
            If tOFNOTIFY.lpOFN <> 0 Then
                CopyMemory tOFN, ByVal tOFNOTIFY.lpOFN, Len(tOFN)
            End If
            'New filter index: tOFN.FilterIndex
'            vFilter = Split(tOFN.Filter, vbNullChar)
'            sTmp = vFilter((tOFN.FilterIndex * 2) - 1)
            ' e.g. "*.OFX"
'            sTmp = Mid$(sTmp, 3)    ' extension
            sTmp = "Bla"
'            tOFN.FileName = sTmp & Space$(1024) & vbNullChar & vbNullChar
        End Select
        OFNHookProc = 0

    Case Else:
        OFNHookProc = 0
    End Select

End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetOutputFileName
' Description:       allow punter to choose output file name
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       31/01/2004-21:27:40
'
' Parameters :       sFileIn (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetOutputFileName(sFilein As String, hOwner As Long, sType As String, iOutType As Long) As String
    Dim OFNW As myOPENFILENAMEW
    Dim OFN As myOPENFILENAME
    Dim sOut As String
    Dim lErr As EDialogError
    Dim i As Long
    Dim xCustomFormats As CustomOutputList
    Dim xCF As CustomOutputFormat
    Dim sTmp As String
    
    DBCSLog sFilein, "Input to GetOutputFileName"
' remove the extension from the output file - then the dialog automatically adds the
' extension corresponding to the filter selection
    sOut = ChangeExtension(sFilein, "")
    
    Set xCustomFormats = New CustomOutputList
    If Not xCustomFormats.Load(Cfg.CustomOutputPath) Then
        LogMessage False, True, "GetOutputFileName: Error loading custom output formats"
    End If
    
    On Error GoTo baleout
    With OFN
        'size of the OFN structure
        .nStructSize = Len(OFN)

        'window owning the dialog
        .hwndOwner = hOwner
      
        'default filename, plus additional padding
        'for the user's final selection(s). Must be
        'double-null terminated
        .FileName = sOut & Space$(1024) & vbNullChar & vbNullChar
        DBCSLog .FileName, "In OPENFILENAME"
        
        'the size of the buffer - in bytes as we are in ANSI mode
        .nMaxFile = LenB(.FileName)
      
        'default extension applied to
        'file if it has no extention
        .DefFileExt = Cfg.OutputFileType & vbNullChar & vbNullChar
                                     
        'space for the file title if a single selection
        'made, double-null terminated, and its size - in bytes as we are in ANSI mode
        .FileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
        .nMaxTitle = LenB(.FileTitle)
      
        'starting folder, double-null terminated
        ' we use a null pointer to force the path in the file name to be used
        .InitialDir = 0&
      
        'the dialog title
        .DialogTitle = LoadResStringL(102) & vbNullChar
    
        sTmp = LoadResStringL(104)
        For Each xCF In xCustomFormats
            sTmp = sTmp & "|" & xCF.FormatName & "|" & xCF.Filters
        Next
        
        .Filter = Replace(sTmp, "|", vbNullChar) & vbNullChar & vbNullChar
        
        If UCase$(Cfg.OutputFileType) = "OFX" Then
            .FilterIndex = 2
        ElseIf UCase$(Cfg.OutputFileType) = "OFC" Then
            .FilterIndex = 1
        ElseIf UCase$(Cfg.OutputFileType) = "QFX" Then
            .FilterIndex = 3
        ElseIf UCase$(Cfg.OutputFileType) = "QIF" Then
            .FilterIndex = 4
        Else
            .FilterIndex = 5
        End If
        .Flags = OFN_HIDEREADONLY + OFN_LONGNAMES + OFN_OVERWRITEPROMPT + OFN_EXPLORER

        DBCSLog .FileName, "File name (" & CStr(.nMaxFile) & " chars) before copymem"
        CopyMemory ByVal VarPtr(OFNW.nStructSize), ByVal VarPtr(OFN.nStructSize), Len(OFNW)
        DBCSLog .FileName, "File name (" & CStr(.nMaxFile) & " chars) after copymem"
        
        OFNW.DialogTitle = StrPtr(OFN.DialogTitle)
        OFNW.FileName = StrPtr(OFN.FileName)
        OFNW.FileTitle = StrPtr(OFN.FileTitle)
        OFNW.Filter = StrPtr(OFN.Filter)
        OFNW.DialogTitle = StrPtr(OFN.DialogTitle)
        OFNW.nMaxTitle = Len(.FileTitle)
        OFNW.nMaxFile = Len(.FileName)
        
        DBCSLog .DialogTitle, "Dialog title"
        DBCSLog .FileName, "File name (" & CStr(.nMaxFile) & " chars)"
        DBCSLog .FileTitle, "File title (" & CStr(.nMaxTitle) & " chars)"
        DBCSLog .Filter, "Filter"
'        DBCSLog .InitialDir, "Initial dir"

        If GetSaveFileName(OFNW) Then
            i = StrLenW(OFNW.FileName)
            sOut = String$(i, vbNullChar)
            CopyMemory ByVal StrPtr(sOut), ByVal OFNW.FileName, i * 2
            Select Case OFNW.FilterIndex
            Case 1: sType = "OFC": iOutType = FileFormatOFC
            Case 2: sType = "OFX": iOutType = FileFormatOFX
            Case 3: sType = "QFX": iOutType = FileFormatQFX
            Case 4: sType = "QIF": iOutType = FileFormatQIF
            Case Else
                sType = xCustomFormats.Item(OFNW.FilterIndex - 4).ScriptName
'                sType = GetDefaultOutputType(sOut)
                iOutType = FileFormatCustom
            End Select
            GetOutputFileName = sOut
        Else
            lErr = CommDlgExtendedError()
            If lErr <> 0 Then
                LogMessage True, True, "Error &H" & Hex(lErr) & " from GetSaveFileName: " & ComDlgErrMsg(lErr)
            End If
            sType = ""
            GetOutputFileName = ""
        End If
    End With
goback:
    DBCSLog GetOutputFileName, "GetOutputFileName returning"
    Exit Function
baleout:
    ShowError "GetOutputFileName"
    Debug.Assert False
    GetOutputFileName = ""
    Resume goback
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetInputFileName
' Description:       allow punter to choose input file name
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       31/01/2004-21:27:40
'
' Parameters :       sFileIn (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetInputFileName(sFilein As String, sInitDir As String, sFilter As String, _
    iFilterIndex As Integer, sTitle As String, hOwner As Long) As String
    Dim OFNW As myOPENFILENAMEW
    Dim OFN As myOPENFILENAME
    Dim sOut As String
    Dim lErr As EDialogError
    Dim i As Long
    
    DBCSLog sFilein, "Input to GetInputFileName"
' remove the extension from the output file - then the dialog automatically adds the
' extension corresponding to the filter selection
    sOut = ChangeExtension(sFilein, "")
    On Error GoTo baleout
    With OFN
        'size of the OFN structure
        .nStructSize = Len(OFN)

        'window owning the dialog
        .hwndOwner = hOwner
      
        'default filename, plus additional padding
        'for the user's final selection(s). Must be
        'double-null terminated
        .FileName = sFilein & Space$(1024) & vbNullChar & vbNullChar
        DBCSLog .FileName, "In OPENFILENAME"
        
        'the size of the buffer - in bytes as we are in ANSI mode
        .nMaxFile = LenB(.FileName)
      
        'default extension applied to
        'file if it has no extention
        .DefFileExt = Cfg.OutputFileType & vbNullChar & vbNullChar
                                     
        'space for the file title if a single selection
        'made, double-null terminated, and its size - in bytes as we are in ANSI mode
        .FileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
        .nMaxTitle = LenB(.FileTitle)
      
        'starting folder, double-null terminated
        ' we use a null pointer to force the path in the file name to be used
        .InitialDir = sInitDir & vbNullChar
      
        'the dialog title
        .DialogTitle = sTitle & vbNullChar
    
        .Filter = Replace(sFilter, "|", vbNullChar) & vbNullChar & vbNullChar
        .FilterIndex = iFilterIndex
        
        .Flags = OFN_LONGNAMES + OFN_OVERWRITEPROMPT + OFN_EXPLORER

        DBCSLog .FileName, "File name (" & CStr(.nMaxFile) & " chars) before copymem"
        CopyMemory ByVal VarPtr(OFNW.nStructSize), ByVal VarPtr(OFN.nStructSize), Len(OFNW)
        DBCSLog .FileName, "File name (" & CStr(.nMaxFile) & " chars) after copymem"
        
        OFNW.DialogTitle = StrPtr(OFN.DialogTitle)
        OFNW.FileName = StrPtr(OFN.FileName)
        OFNW.InitialDir = StrPtr(OFN.InitialDir)
        OFNW.FileTitle = StrPtr(OFN.FileTitle)
        OFNW.Filter = StrPtr(OFN.Filter)
        OFNW.DialogTitle = StrPtr(OFN.DialogTitle)
        OFNW.nMaxTitle = Len(.FileTitle)
        OFNW.nMaxFile = Len(.FileName)
        
        DBCSLog .DialogTitle, "Dialog title"
        DBCSLog .FileName, "File name (" & CStr(.nMaxFile) & " chars)"
        DBCSLog .FileTitle, "File title (" & CStr(.nMaxTitle) & " chars)"
        DBCSLog .Filter, "Filter"
'        DBCSLog .InitialDir, "Initial dir"

        If GetOpenFileName(OFNW) Then
            i = StrLenW(OFNW.FileName)
            sOut = String$(i, vbNullChar)
            CopyMemory ByVal StrPtr(sOut), ByVal OFNW.FileName, i * 2
            GetInputFileName = sOut
'            GetOutputFileName = Left$(.FileName, InStr(.FileName, vbNullChar) - 1)
        Else
            lErr = CommDlgExtendedError()
            If lErr <> 0 Then
                LogMessage True, True, "Error &H" & Hex(lErr) & " from GetOpenFileName: " & ComDlgErrMsg(lErr)
            End If
            GetInputFileName = ""
        End If
    End With
goback:
    DBCSLog GetInputFileName, "GetInputFileName returning"
    Exit Function
baleout:
    ShowError "GetInputFileName"
    Debug.Assert False
    GetInputFileName = ""
    Resume goback
End Function

Public Function ComDlgErrMsg(lErr As Long) As String
    Dim s As String
    Select Case lErr
    Case CDERR_DIALOGFAILURE: s = "The dialog box could not be created."

    Case CDERR_STRUCTSIZE: s = "The lStructSize member of the initialization structure for the corresponding common dialog box is invalid."
    Case CDERR_INITIALIZATION: s = "The common dialog box function failed during initialization. "
    Case CDERR_NOTEMPLATE: s = "The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding template."
    Case CDERR_NOHINSTANCE: s = "The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding instance handle."
    Case CDERR_LOADSTRFAILURE: s = "The common dialog box function failed to load a specified string."
    Case CDERR_FINDRESFAILURE: s = "The common dialog box function failed to find a specified resource."
    Case CDERR_LOADRESFAILURE: s = "The common dialog box function failed to load a specified resource."
    Case CDERR_LOCKRESFAILURE: s = "The common dialog box function failed to lock a specified resource."
    Case CDERR_MEMALLOCFAILURE: s = "The common dialog box function was unable to allocate memory for internal structures."
    Case CDERR_MEMLOCKFAILURE: s = "The common dialog box function was unable to lock the memory associated with a handle."
    Case CDERR_NOHOOK: s = "The ENABLEHOOK flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a pointer to a corresponding hook procedure."
    Case CDERR_REGISTERMSGFAIL: s = "The RegisterWindowMessage function returned an error code when it was called by the common dialog box function."
#If False Then
    Case PDERR_SETUPFAILURE: s = ""
    Case PDERR_PARSEFAILURE: s = ""
    Case PDERR_RETDEFFAILURE: s = ""
    Case PDERR_LOADDRVFAILURE: s = ""
    Case PDERR_GETDEVMODEFAIL: s = ""
    Case PDERR_INITFAILURE: s = ""
    Case PDERR_NODEVICES: s = ""
    Case PDERR_NODEFAULTPRN: s = ""
    Case PDERR_DNDMMISMATCH: s = ""
    Case PDERR_CREATEICFAILURE: s = ""
    Case PDERR_PRINTERNOTFOUND: s = ""
    Case PDERR_DEFAULTDIFFERENT: s = ""

    Case CFERR_NOFONTS: s = ""
    Case CFERR_MAXLESSTHANMIN: s = ""
#End If
    Case FNERR_SUBCLASSFAILURE: s = "An attempt to subclass a list box failed because sufficient memory was not available."
    Case FNERR_INVALIDFILENAME: s = "A file name is invalid."
    Case FNERR_BUFFERTOOSMALL: s = "The buffer pointed to by the lpstrFile member of the OPENFILENAME structure is too small for the file name specified by the user. The first two bytes of the lpstrFile buffer contain an integer value specifying the size, in TCHARs, required to receive the full name."

    Case Else
        s = "Common Dialog Error &H" & Hex$(lErr)
    End Select
    ComDlgErrMsg = s
End Function
