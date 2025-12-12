Attribute VB_Name = "ComDialog"
Option Explicit
' $Revision: 1 $
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
  InitialDir        As Long
  DialogTitle       As String
  Flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  DefFileExt        As String
  nCustData         As Long
  fnHook            As Long
  sTemplateName     As String
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

Public Declare Function GetOpenFileName Lib "comdlg32" _
    Alias "GetOpenFileNameA" _
   (pOpenfilename As myOPENFILENAME) As Long

Public Declare Function GetSaveFileName Lib "comdlg32" _
   Alias "GetSaveFileNameA" _
  (pOpenfilename As myOPENFILENAME) As Long

Public Declare Function GetShortPathName Lib "kernel32" _
    Alias "GetShortPathNameA" _
   (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)



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
Public Function OFNHookProc(ByVal hwnd As Long, _
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
Public Function GetOutputFileName(sFilein As String, frmOwner As Form) As String
    Dim OFN As myOPENFILENAME
    Dim sOut As String
' remove the extension from the output file - then the dialog automatically adds the
' extension corresponding to the filter selection
    sOut = ChangeExtension(sFilein, "")
    On Error GoTo baleout
    With OFN
        'size of the OFN structure
        .nStructSize = Len(OFN)

        'window owning the dialog
        .hwndOwner = frmOwner.hwnd
      
        'default filename, plus additional padding
        'for the user's final selection(s). Must be
        'double-null terminated
        .FileName = sOut & Space$(1024) & vbNullChar & vbNullChar
      
        'the size of the buffer
        .nMaxFile = Len(.FileName)
      
        'default extension applied to
        'file if it has no extention
        .DefFileExt = Cfg.OutputFileType & vbNullChar & vbNullChar
                                     
        'space for the file title if a single selection
        'made, double-null terminated, and its size
        .FileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
        .nMaxTitle = Len(.FileTitle)
      
        'starting folder, double-null terminated
        ' we use a null pointer to force the path in the file name to be used
        .InitialDir = 0&
      
        'the dialog title
        .DialogTitle = LoadResStringL(102)
    
        .Filter = Replace(LoadResStringL(104), "|", vbNullChar) & vbNullChar & vbNullChar
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

        If GetSaveFileName(OFN) Then
            GetOutputFileName = Left$(.FileName, InStr(.FileName, vbNullChar) - 1)
        Else
            GetOutputFileName = ""
        End If
    End With
goback:
    Exit Function
baleout:
'    MsgBox "An error has occurred"
    Debug.Assert False
    GetOutputFileName = ""
    Resume goback
End Function

