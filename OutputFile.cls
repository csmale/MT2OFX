VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OutputFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : OutputFile
'    Project    : MT2OFX
'
'    Description: Output Stream with code page handling
'
'    Modified   : $Author: Colin $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/OutputFile.cls 6     15/11/10 0:10 Colin $"
' $History: OutputFile.cls $
' 
' *****************  Version 6  *****************
' User: Colin        Date: 15/11/10   Time: 0:10
' Updated in $/MT2OFX
'
' *****************  Version 5  *****************
' User: Colin        Date: 30/08/09   Time: 13:16
' Updated in $/MT2OFX
' copied some declarations here to improve decoupling
' now defaults to producing BOM for UTF8
'
' *****************  Version 3  *****************
' User: Colin        Date: 19/04/08   Time: 23:21
' Updated in $/MT2OFX
'
' *****************  Version 2  *****************
' User: Colin        Date: 18/10/07   Time: 23:14
' Updated in $/MT2OFX

'</CSCC>

Private Const CP_ACP = 0
Private Const CP_UNKNOWN = -1
Private Const CP_UTF7 = 65000
Private Const CP_UTF8 = 65001
Private Const CP_UNICODELITTLE = 1200
Private Const CP_UNICODEBIG = 1201
Private Declare Function GetACP Lib "kernel32" () As Long

Private Declare Function WideCharToMultiByte Lib "unicows.dll" (ByVal CodePage As Long, _
    ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, _
    lpUsedDefaultChar As Long) As Long

Private Const WC_NO_BEST_FIT_CHARS As Long = &H400

Private xiCodePage As Long
Private xsFileName As String
Private xsLines() As String
Private xsLineEnd As String
Private xiFile As Integer
Public OutputBOM As Boolean

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       LineEnd
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/18/2007-23:09:18
'
' Parameters :       s (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Let LineEnd(s As String)
    xsLineEnd = s
End Property

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       LineEnd
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/18/2007-23:09:18
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get LineEnd() As String
    LineEnd = xsLineEnd
End Property

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CodePage
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/18/2007-23:09:18
'
' Parameters :       cp (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Let CodePage(cp As Long)
    xiCodePage = cp
End Property

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CodePage
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/18/2007-23:09:18
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
' Procedure  :       OpenFile
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/18/2007-23:09:18
'
' Parameters :       sFile (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function OpenFile(sFile As String) As Boolean
    On Error GoTo baleout
    
    OpenFile = False
    If xiFile > 0 Then
        Err.Raise vbObjectError + 1000, "OutputFile.OpenFile", "OutputFile already has a file open"
        Exit Function
    End If
    If Dir(sFile) <> "" Then Kill sFile
    xiFile = FreeFile
    Open sFile For Binary Access Write As #xiFile
    xsFileName = sFile
    If xiCodePage = CP_ACP Then
        xiCodePage = GetACP()
    End If
    OpenFile = True
goback:
    Exit Function
baleout:
    OpenFile = False
    xiFile = 0
    LogMessage True, True, Err.Description, "Error opening output file", vbOKOnly + vbCritical
    Resume goback
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CloseFile
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/18/2007-23:09:18
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function CloseFile() As Boolean
    Dim sTmp As String
    Dim lFlags As Long
    Dim bData() As Byte
    Dim textLen As Long
    Dim bOutput() As Byte
    Dim tmpLen As Long
    
    Debug.Assert (xiFile > 0)
    If xiFile = 0 Then
        CloseFile = False
        Exit Function
    End If

' emit any byte order marks
    If OutputBOM Then
        Select Case xiCodePage
        Case CP_UTF8
            ReDim bOutput(2): bOutput(0) = &HEF: bOutput(1) = &HBB: bOutput(2) = &HBF
            Put #xiFile, , bOutput()
        Case CP_UNICODEBIG
            ReDim bOutput(1): bOutput(0) = &HFE: bOutput(1) = &HFF
            Put #xiFile, , bOutput()
        Case CP_UNICODELITTLE
            ReDim bOutput(1): bOutput(0) = &HFF: bOutput(1) = &HFE
            Put #xiFile, , bOutput()
        End Select
    End If

' make it into one big string
    sTmp = Join(xsLines, xsLineEnd)
' convert the string (unicode) into a byte array - 2 bytes per char
    bData = sTmp
' make a buffer for the converted data
    tmpLen = 0
    ReDim bOutput(0)
    textLen = Len(sTmp)
    tmpLen = WideCharToMultiByte(CLng(xiCodePage), lFlags, _
        ByVal VarPtr(bData(0)), textLen, _
        ByVal VarPtr(bOutput(0)), tmpLen, _
        ByVal 0&, ByVal 0&)
    If tmpLen = 0 Then Exit Function
    ReDim bOutput(0 To tmpLen - 1)
' now we are ready to convert the byte array into the desired code page
    tmpLen = WideCharToMultiByte(CLng(xiCodePage), lFlags, _
        ByVal VarPtr(bData(0)), textLen, _
        ByVal VarPtr(bOutput(0)), tmpLen, _
        ByVal 0&, ByVal 0&)
    If tmpLen = 0 Then Exit Function

    Put #xiFile, , bOutput()

    Close #xiFile
    xiFile = 0
    CloseFile = True
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       PrintLine
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/18/2007-23:09:18
'
' Parameters :       sLine (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub PrintLine(sLine As String)
    ReDim Preserve xsLines(UBound(xsLines) + 1)
    xsLines(UBound(xsLines) - 1) = sLine
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CharInCodepage
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/18/2007-23:09:18
'
' Parameters :       c (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function CharInCodepage(c As String) As Boolean
    Dim sTmp As String
    Dim lFlags As Long
    Dim bData() As Byte
    Dim textLen As Long
    Dim bOutput(4) As Byte
    Dim tmpLen As Long
    Dim iUsedDefault As Long

    CharInCodepage = False
' unicode code pages can handle anything
    Select Case xiCodePage
    Case CP_UTF7, CP_UTF8, CP_UNICODEBIG, CP_UNICODELITTLE
        CharInCodepage = True
        Exit Function
    End Select
' empty string is always ok (shouldn't really happen)
    If Len(c) < 1 Then
        Debug.Assert False
        CharInCodepage = True
        Exit Function
    End If
' convert the string (unicode) into a byte array - 2 bytes per char
    bData = Left(c, 1)
    textLen = 1
' make a buffer for the converted data - 4 bytes per char is as much as we will ever need
    tmpLen = 4
' now we are ready to convert the byte array into the desired code page
' flag specifies that the api must flag the usage of the "default char" which will be the case when
' the character cannot be represented in the selected code page.
    lFlags = WC_NO_BEST_FIT_CHARS
    tmpLen = WideCharToMultiByte(CLng(xiCodePage), lFlags, _
        ByVal VarPtr(bData(0)), textLen, _
        ByVal VarPtr(bOutput(0)), tmpLen, _
        ByVal 0&, iUsedDefault)
        
    If tmpLen = 0 Then Exit Function
    CharInCodepage = (iUsedDefault = 0)
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Class_Initialize
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/18/2007-23:09:18
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Class_Initialize()
    xiFile = 0
    xsLineEnd = vbCrLf
    xiCodePage = GetACP()
    ReDim xsLines(0)
    OutputBOM = True
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Class_Terminate
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       10/18/2007-23:09:18
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Class_Terminate()
    If xiFile > 0 Then
        Close #xiFile
        xiFile = 0
    End If
End Sub

