Attribute VB_Name = "modInIDE"
Option Explicit

' $Header: /MT2OFX/modInIDE.bas 6     15/06/09 19:25 Colin $

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Dim m_bInDevelopment As Boolean

' get window word constants
Const GWW_HWNDPARENT = (-8)

' Sub Form_Load()
    ' note: declare gDebugMode in any module --> Global gDebugMode as Integer
'    gDebugMode = InIDE(hWnd)

'    If gDebugMode Then
'        Debug.Print "IDE detected"
'    End If

'End Sub

' ------------------------------------------------------------------------------------
' ROUTINE: InIDE:BOOL, Params( inhwnd InputOnly )
' Purpose: to determine if the program is running in the Integrated Development Environment
'
' Description: Uses the class of the hidden parent window to determine if the program
' is running in the IDE or is compiled into an EXE.
'
' INPUT: inhwnd -- the window handle of the calling window
' OUTPUT: return code is True if the program is running in the IDE, False if it is an EXE
'
' ------------------------------------------------------------------------------------
'
Function InIDE(ByVal inhwnd As Long) As Boolean

Dim parent As Long, pclass As String, nlen As Long

    parent = GetWindowLong(inhwnd, GWW_HWNDPARENT)
    pclass = Space$(32)
    nlen = GetClassName(parent, pclass, 31)
    pclass = Left$(pclass, nlen)

    If InStr(pclass, "RT") Then
        InIDE = False
    Else
        InIDE = True
    End If
End Function

Public Function InDevelopment() As Boolean
   ' Debug.Assert code not run in an EXE. Therefore
   ' m_bInDevelopment variable is never set.
   Debug.Assert InDevelopmentHack() = True
   InDevelopment = m_bInDevelopment
End Function

Private Function InDevelopmentHack() As Boolean
   ' .... '
   m_bInDevelopment = True
   InDevelopmentHack = m_bInDevelopment
End Function

