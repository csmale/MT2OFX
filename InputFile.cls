VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InputFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : InputFile
'    Project    : MT2OFX
'
'    Description: Presents standardised interface to read a text file
'                 Handles code pages, UTF-8/UTF-16, CR/LF differences
'
'    Modified   : $Author: Colin $ $Date: 15/11/10 0:09 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/InputFile.cls 8     15/11/10 0:09 Colin $"
' $History: InputFile.cls $
' 
' *****************  Version 8  *****************
' User: Colin        Date: 15/11/10   Time: 0:09
' Updated in $/MT2OFX
'
' *****************  Version 7  *****************
' User: Colin        Date: 30/08/09   Time: 13:15
' Updated in $/MT2OFX
' copied some declarations here to improve decoupling
'
' *****************  Version 6  *****************
' User: Colin        Date: 15/06/09   Time: 19:25
' Updated in $/MT2OFX
' For transfer to new laptop
'
' *****************  Version 5  *****************
' User: Colin        Date: 25/11/08   Time: 22:23
' Updated in $/MT2OFX
' moving vss server!
'
' *****************  Version 3  *****************
' User: Colin        Date: 19/04/08   Time: 23:25
' Updated in $/MT2OFX
'
' *****************  Version 2  *****************
' User: Colin        Date: 18/10/07   Time: 23:14
' Updated in $/MT2OFX

'</CSCC>

Private Const CP_UNKNOWN = -1
Private Const CP_UTF8 = 65001
Private Const CP_UNICODELITTLE = 1200
Private Const CP_UNICODEBIG = 1201

Private Declare Function MultiByteToWideChar Lib "unicows.dll" ( _
    ByVal CodePage As Long, ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

Private Declare Function WideCharToMultiByte Lib "unicows.dll" (ByVal CodePage As Long, _
    ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, _
    lpUsedDefaultChar As Long) As Long

Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetACP Lib "kernel32" () As Long

Private xiCodePage As Long
Private xiPos As Long
Private xiLength As Long
Private xiLengthB As Long
Private xsFileName As String
Private bData() As Byte
Private sLines() As String
Private nLines As Long
Private iLine As Long

Private xbConvertPending As Boolean

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CodePage (Property Get)
' Description:       Returns current code page
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/9/2007-23:51:04
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get CodePage() As Long
    CodePage = xiCodePage
End Property

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CodePage (Property Let)
' Description:       Changes the code page used to interpret the file
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/9/2007-23:51:04
'
' Parameters :       i (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Let CodePage(cp As Long)
    xiCodePage = cp
    If xiCodePage = CP_UNKNOWN Then xiCodePage = GetACP()
    xbConvertPending = True
End Property

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       FileName
' Description:       Returns the name of the opened file
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/10/2007-19:34:25
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get FileName() As String
    FileName = xsFileName
End Property
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Pos (Property Get)
' Description:       Returns the current byte position
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/9/2007-23:51:04
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get Pos() As Long
    Pos = xiPos
End Property

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Pos (Property Let)
' Description:       Sets current byte position
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/9/2007-23:51:04
'
' Parameters :       iNew (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Let Pos(iNew As Long)
    xiPos = 0
    iLine = 0
    Do While iLine < nLines
        If xiPos + Len(sLines(iLine)) + 1 > iNew Then
            Exit Do
        End If
        xiPos = xiPos + Len(sLines(iLine)) + 1
        iLine = iLine + 1
    Loop
End Property

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Length
' Description:       Returns length of file in characters
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/9/2007-23:51:04
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get Length() As Long
    Call DoConvert
    Length = xiLength
End Property

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       LengthB
' Description:       Returns length of file in bytes
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/9/2007-23:51:04
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get LengthB() As Long
    LengthB = xiLengthB
End Property

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       AtEOF
' Description:       Returns EOF indication
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/9/2007-23:51:04
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get AtEOF() As Boolean
    Call DoConvert
    AtEOF = (iLine > nLines)
End Property

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       EntireFile
' Description:       Returns entire file as a string, joined with CR
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/9/2007-23:51:04
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get EntireFile() As String
    Dim sTmp As String
    Call DoConvert
    sTmp = Join(sLines, vbCr)
    EntireFile = sTmp
End Property

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Rewind
' Description:       Resets file to the beginning
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/9/2007-23:51:04
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub Rewind()
    iLine = 0
    xiPos = 0
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OpenFile
' Description:       Opens a file and reads it in
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/9/2007-23:51:03
'
' Parameters :       sFile (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function OpenFile(sFile As String) As Boolean
    Dim iFile As Integer
    Dim iLen As Long
    
    On Error GoTo baleout
    
    nLines = -1
    iLine = 0
' Open for Binary will create a file if it doesn't exist, even with Access Read!
    iFile = FreeFile
    Open sFile For Input Access Read As #iFile
' if we get here the file must exist, so we can reopen it in binary mode
    Close #iFile
    Open sFile For Binary Access Read As #iFile
    xiLengthB = LOF(iFile)
    ReDim bData(xiLengthB)
    Get #iFile, , bData()
    Close #iFile
    
    xsFileName = sFile
    xbConvertPending = True
    OpenFile = True
goback:
    Exit Function
baleout:
    OpenFile = False
    Err.Raise Err.Number, "InputFile.OpenFile", Err.Description
'    Resume goback
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CloseFile
' Description:       Forgets data about the file
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/9/2007-23:51:03
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub CloseFile()
    nLines = -1
    iLine = 0
    ReDim sLines(0)
    xsFileName = ""
    xbConvertPending = False
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ReadLine
' Description:       Returns the next line of text and advances the pointer
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/9/2007-23:51:03
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ReadLine() As String
    Dim sTmp As String
    Call DoConvert
    If iLine <= nLines Then
        sTmp = sLines(iLine)
        iLine = iLine + 1
    Else
        sTmp = ""
    End If
    xiPos = xiPos + Len(sTmp) + 1
    ReadLine = sTmp
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoConvert
' Description:       Converts raw byte array to an array of strings, using the appropriate code page
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/9/2007-23:51:03
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DoConvert()
    Dim sTmp As String
    Dim tmpLen As Long
    Dim lFlags As Long
    Dim tmpArr() As Byte
    Dim i As Long
    Dim iCP As Long
    
    If Not xbConvertPending Then Exit Sub
    xbConvertPending = False
    If xiLengthB = 0 Then Exit Sub
    
    tmpLen = xiLengthB * 2
    ReDim Preserve tmpArr(tmpLen)

' special treatment for unicode utf-8/utf-16 files
    iCP = xiCodePage
    If bData(0) = &HEF And bData(1) = &HBB And bData(2) = &HBF Then
        iCP = CP_UTF8
    ElseIf bData(0) = &HFF And bData(1) = &HFE Then
        iCP = CP_UNICODELITTLE
    ElseIf bData(0) = &HFE And bData(1) = &HFF Then
        iCP = CP_UNICODEBIG
    End If
    If iCP = CP_UNKNOWN Then iCP = GetACP()
    xiCodePage = iCP

' get the new string to tmpArr
    If iCP <> CP_UNICODEBIG And iCP <> CP_UNICODELITTLE Then
        tmpLen = MultiByteToWideChar(CLng(iCP), lFlags, ByVal VarPtr(bData(0)), -1, ByVal VarPtr(tmpArr(0)), tmpLen)

' check for conversion errors
        If tmpLen = 0 Then
            Err.Raise GetLastError(), "MultiByteToWideChar", "Input Code Page Conversion Error"
            Exit Sub
        End If
    ElseIf iCP = CP_UNICODELITTLE Then
        For i = 0 To xiLengthB - 2
            tmpArr(i) = bData(i + 2)
        Next
    ElseIf iCP = CP_UNICODEBIG Then
        For i = 0 To xiLengthB - 2 Step 2
            tmpArr(i) = bData(i + 3)
            tmpArr(i + 1) = bData(i + 2)
        Next
    End If

' convert unicode bytes array to unicode string
    sTmp = tmpArr()
    i = InStr(sTmp, vbNullChar)
    If i > 0 Then  ' make sure we blow away any terminating null
        tmpLen = i - 1
    End If
    sTmp = Left$(sTmp, tmpLen)
' we may have a BOM to skip
    If AscW(Left$(sTmp, 1)) = &HFEFF Then
        sTmp = Mid$(sTmp, 2)
    End If
    sTmp = NormaliseLineEndings(sTmp)   ' turns everything to a CR
    sLines = Split(sTmp, vbCr)

' work out length of text in lines and characters
    nLines = UBound(sLines)
    xiLength = 0
    For i = 0 To nLines
        xiLength = xiLength + Len(sLines(i)) + 1
    Next
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Class_Initialize
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/9/2007-23:51:03
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Class_Initialize()
    iLine = -1
    nLines = -1
    ReDim sLines(0)
    xiCodePage = GetACP()
End Sub


