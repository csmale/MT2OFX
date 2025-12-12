Attribute VB_Name = "GUID"
Option Explicit

' $Header: /MT2OFX/GUID.bas 3     1/10/04 21:21 Colin $

' data type used by CoCreateGuid and StringFromGUID2
Public Type GUID_t
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

' prototype API functions
Private Declare Function CoCreateGuid Lib "OLE32.dll" (pGuid As GUID_t) As Long

Private Declare Function StringFromGUID2 Lib "OLE32.dll" _
    (ByRef rguid As GUID_t, ByVal lpsz As String, ByVal cchMax As Long) As Integer

' Here is the main function definition:

' our function: returns a string GUID
Function GenerateGUID() As String

Const S_OK = 0

Dim NewGUID As GUID_t
Dim strNewGuid As String
Dim iChars As Integer, lReturn As Long

strNewGuid = Space(100)

lReturn = CoCreateGuid(NewGUID)

If (lReturn <> S_OK) Then
    MyMsgBox "CoCreateGuid failed!" & vbNewLine & vbNewLine & _
           "(It's not your fault.)", vbExclamation + vbOKOnly
    Exit Function
End If

' convert binary GUID to string form
iChars = StringFromGUID2(NewGUID, strNewGuid, Len(strNewGuid))
' convert string to ANSI
strNewGuid = StrConv(strNewGuid, vbFromUnicode)
' remove trailing null character
strNewGuid = Left(strNewGuid, iChars - 1)
' MSI likes only uppercase letters in GUID
strNewGuid = UCase(strNewGuid)

GenerateGUID = strNewGuid

End Function


