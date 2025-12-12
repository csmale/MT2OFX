Attribute VB_Name = "Rubbish"
Option Explicit

Private Function ByteArrayToString(b() As Byte) As String
    Dim i As Long
    Dim c As Long
    On Error GoTo bad
    Dim s As String
    For i = LBound(b) To UBound(b) - 1 Step 2
        c = (b(i) + (CLng(256) * CLng(b(i + 1)))) And &HFFFF
If c > 255 Then
'    MsgBox "found wide char &H" & Hex(c)
    If (CLng(AscW(ChrW(c))) And &HFFFF&) <> c Then
        MsgBox "corruption: becomes &H" & (AscW(ChrW(c)) And &HFFFF&)
    End If
End If
        s = s & ChrW(c)
    Next
    ByteArrayToString = s
    Exit Function
bad:
    MsgBox "err at byte " & CStr(i) & " = &H" & Hex(c) & "(" & CStr(b(i)) & "," & CStr(b(i + 1)) & "): " & Err.Description
End Function

