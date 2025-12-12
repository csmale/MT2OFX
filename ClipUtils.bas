Attribute VB_Name = "ClipUtils"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : ClipUtils
'    Project    : MT2OFX
'
'    Description: Clipboard access functions
'
'    Modified   : $Author: Colin $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/ClipUtils.bas 9     15/06/09 19:24 Colin $"
' $History: ClipUtils.bas $
' 
' *****************  Version 9  *****************
' User: Colin        Date: 15/06/09   Time: 19:24
' Updated in $/MT2OFX
' For transfer to new laptop
'
' *****************  Version 8  *****************
' User: Colin        Date: 25/11/08   Time: 22:14
' Updated in $/MT2OFX
' moving vss server!

'</CSCC>

Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) _
   As Long
Private Declare Function GlobalAlloc Lib "kernel32" ( _
   ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function SetClipboardData Lib "user32" ( _
   ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" _
    (ByVal lpString As String) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) _
   As Long
Private Declare Function GlobalUnlock Lib "kernel32" ( _
   ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
   pDest As Any, pSource As Any, ByVal cbLength As Long)
Private Declare Function GetClipboardData Lib "unicows.dll" ( _
   ByVal wFormat As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "unicows.dll" (wFormat As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" ( _
   ByVal lpData As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Private Const m_sDescription = _
                  "Version:1.0" & vbCrLf & _
                  "StartHTML:aaaaaaaaaa" & vbCrLf & _
                  "EndHTML:bbbbbbbbbb" & vbCrLf & _
                  "StartFragment:cccccccccc" & vbCrLf & _
                  "EndFragment:dddddddddd" & vbCrLf

Private Const CF_UNICODETEXT As Long = 13
Private Const CF_LOCALE As Long = 16

Private Declare Function MultiByteToWideChar Lib "unicows.dll" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long) As Long

Private m_cfHTMLClipFormat As Long

Private Function RegisterCF() As Long
   'Register the HTML clipboard format
   If (m_cfHTMLClipFormat = 0) Then
      m_cfHTMLClipFormat = RegisterClipboardFormat("HTML Format")
   End If
   RegisterCF = m_cfHTMLClipFormat
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ClipTest
' Procedure  :       GetHTMLClipboard
' Description:

' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       02/09/2004-16:29:36
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetHTMLClipboard(FullHTML As Boolean) As String
   Dim sData As String
   Dim sData2 As String
   Dim bData() As Byte
   Dim sStartHeader As String
   Dim sEndHeader As String
   
    If FullHTML Then
        sStartHeader = "StartHTML:"
        sEndHeader = "EndHTML:"
    Else
        sStartHeader = "StartFragment:"
        sEndHeader = "EndFragment:"
    End If
   
   If RegisterCF = 0 Then Exit Function
   
   If CBool(OpenClipboard(0)) Then
   
      Dim hMemHandle As Long, lpData As Long
      Dim nClipSize As Long
      
      GlobalUnlock hMemHandle

      'Retrieve the data from the clipboard
      hMemHandle = GetClipboardData(m_cfHTMLClipFormat)
      
      If CBool(hMemHandle) Then
               
         lpData = GlobalLock(hMemHandle)
         If lpData <> 0 Then
            nClipSize = lstrlen(lpData)
            sData = String(nClipSize + 10, 0)
            ReDim bData(nClipSize + 10)

            Call CopyMemory(ByVal sData, ByVal lpData, nClipSize)
            Call CopyMemory(bData(0), ByVal lpData, nClipSize)
            
            Dim nStartFrag As Long, nEndFrag As Long
            Dim nIndx As Long
            
            'If StartFragment appears in the data's description,
            'then retrieve the offset specified in the description
            'for the start of the fragment. Likewise, if EndFragment
            'appears in the description, then retrieve the
            'corresponding offset.
            nIndx = InStr(sData, sStartHeader)
            If nIndx Then
               nStartFrag = CLng(Mid(sData, _
                                 nIndx + Len(sStartHeader), 10))

            End If
            nIndx = InStr(sData, sEndHeader)
            If nIndx Then
               nEndFrag = CLng(Mid(sData, nIndx + Len(sEndHeader), 10))
            End If
            
            'Return the fragment given the starting and ending
            'offsets
            If (nStartFrag > 0 And nEndFrag > 0) Then
                sData2 = GetUTF8Data(bData, nStartFrag, nEndFrag)
'               sData = Mid(sData, nStartFrag + 1, _
'                                 (nEndFrag - nStartFrag))
'                sData2 = UTF8ToString(sData)
'            MsgBox sData2
                GetHTMLClipboard = sData2
            
            End If
                        
         End If
      
      End If

   
      Call CloseClipboard
   End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ClipTest
' Procedure  :       UTF8ToString
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       02/09/2004-17:28:04
'
' Parameters :       sIn (String)
'--------------------------------------------------------------------------------
'</CSCM>
Function GetUTF8Data(bData() As Byte, iStart As Long, iEnd As Long) As String
    Dim sTmp As String
    Dim tmpLen As Long
    Dim lFlags As Long
    Dim tmpArr() As Byte
    Dim i As Long
    Dim iCP As Long
    
    tmpLen = (iEnd - iStart + 1) * 2
    ReDim tmpArr(tmpLen)

' special treatment for unicode utf-8/utf-16 files
    iCP = CP_UTF8
    
' get the new string to tmpArr
    tmpLen = MultiByteToWideChar(CLng(iCP), lFlags, ByVal VarPtr(bData(iStart)), (iEnd - iStart + 1), ByVal VarPtr(tmpArr(0)), tmpLen)

' check for conversion errors
    If tmpLen = 0 Then
        Err.Raise GetLastError(), "MultiByteToWideChar", "Input Code Page Conversion Error"
        Exit Function
    End If

' convert unicode bytes array to unicode string
    sTmp = tmpArr()
    sTmp = NormaliseLineEndings(Left$(sTmp, tmpLen - 1))  ' don't bother with terminating null and standardise on CR

    GetUTF8Data = sTmp
End Function



'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CaptureTextClipboard
' Description:       get text clipboard into a file and return its name
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       26 Nov 2004-23:00:33
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function CaptureTextClipboard() As String
    Dim sClip As String
    Dim iClip As Integer
    Dim sIn As String
    Dim xTemp As New OutputFile
    
    On Error GoTo baleout
    CaptureTextClipboard = ""
    sClip = GetUnicodeClipboard()
    If Len(sClip) = 0 Then
        sClip = Clipboard.GetText
    End If
    If Len(sClip) = 0 Then
        Exit Function
    End If
    sIn = GetTempFile("TMP")    ' cannot assume the extension!
' save the clipboard as UTF8 so we preserve all unicode chars
    xTemp.CodePage = CP_UTF8
    If xTemp.OpenFile(sIn) Then
        xTemp.OutputBOM = True
        xTemp.PrintLine sClip
        xTemp.CloseFile
    End If
    CaptureTextClipboard = sIn
goback:
    Exit Function
baleout:
    ShowError "CaptureTextClipboard"
    CaptureTextClipboard = ""
    Resume goback
End Function

Public Function GetUnicodeClipboard() As String
    Dim bData() As Byte
    Dim hMemHandle As Long, lpData As Long
    Dim nClipSize As Long

    If CBool(OpenClipboard(0)) Then
        GlobalUnlock hMemHandle
        'Retrieve the data from the clipboard
        hMemHandle = GetClipboardData(CF_UNICODETEXT)
        If CBool(hMemHandle) Then
            lpData = GlobalLock(hMemHandle)
            If lpData <> 0 Then
                nClipSize = lstrlenW(lpData)
                ReDim bData(nClipSize * 2)
                Call CopyMemory(bData(0), ByVal lpData, nClipSize * 2)
                GetUnicodeClipboard = bData()
            End If
        End If
    End If

    Call CloseClipboard
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetClipboardLocale
' Description:       Returns locale of the clipboard text (CF_TEXT) as an LCID
'                    e.g. English (US) is 0x0809
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       5/5/2009-18:18:33
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetClipboardLocale() As Long
    Dim bData() As Byte
    Dim hMemHandle As Long, lpData As Long
    Dim nClipSize As Long

    If CBool(OpenClipboard(0)) Then
        GlobalUnlock hMemHandle
        'Retrieve the data from the clipboard
        hMemHandle = GetClipboardData(CF_LOCALE)
        If CBool(hMemHandle) Then
            lpData = GlobalLock(hMemHandle)
            If lpData <> 0 Then
                nClipSize = lstrlenW(lpData)
                ReDim bData(nClipSize * 2)
                Call CopyMemory(bData(0), ByVal lpData, nClipSize * 2)
                GetClipboardLocale = (bData(1) * 256) + bData(0)
            End If
        End If
    End If

    Call CloseClipboard

End Function

