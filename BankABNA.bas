Attribute VB_Name = "BankABNA"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : BankABNA
'    Project    : MT2OFX
'
'    Description: ABN Amro embedded processing
'
'    Modified   : $Author: Colin $ $Date: 6/03/05 23:41 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/BankABNA.bas 3     6/03/05 23:41 Colin $"
' $History: BankABNA.bas $
' 
' *****************  Version 3  *****************
' User: Colin        Date: 6/03/05    Time: 23:41
' Updated in $/MT2OFX
'</CSCC>

Const IniSectionSpecialPayeeNames = "SpecialPayeeNames"
Const IniPatternPrefix = "Pattern"
Const IniPayeePrefix = "Payee"

Public Function ABNA_FindPayee(t As Txn) As String
    Dim v As Variant
    Dim sPayee As String
    Dim iTmp As Integer
    Dim sMemo As String
    
    sMemo = t.FurtherInfo
    If sMemo = "" Then
        ABNA_FindPayee = sMemo
        Exit Function
    End If
    sPayee = GetSpecialPayee(sMemo)
    If sPayee <> "" Then
        ABNA_FindPayee = sPayee
        Exit Function
    End If
    v = Split(sMemo, Cfg.MemoDelimiter)
    If Not IsArray(v) Then
        ABNA_FindPayee = sMemo
        Exit Function
    End If
    sPayee = v(0)
    If Left$(sPayee, 16) = "PROV.  TELEGIRO " Then
        sPayee = Mid$(sPayee, 17)
    End If
    If Left$(sPayee, 14) = "BETAALAUTOMAAT" Then
        sPayee = v(1)
        iTmp = InStr(sPayee, ",")
        If iTmp <> 0 Then
            sPayee = Trim$(Left$(sPayee, iTmp - 1))
        End If
    ElseIf IsNumeric(Left$(sPayee, 1)) Then ' starts with account number
        sPayee = Trim$(Mid$(sPayee, 14))
        If sPayee = "" Then sPayee = v(1)
    ElseIf Left$(sPayee, 5) = "GIRO " Then
        sPayee = Mid$(sPayee, 6)
        While Len(sPayee) > 0 And (Left$(sPayee, 1) = " " Or IsNumeric(Left$(sPayee, 1)))
            sPayee = Mid$(sPayee, 2)
        Wend
        sPayee = Trim$(sPayee)
        If sPayee = "" Then
            sPayee = v(1)
        End If
    ElseIf Left$(sPayee, 2) = "NI" And IsNumeric(Mid$(sPayee, 3, 1)) Then
        sPayee = v(1)
    ElseIf Left$(sPayee, 6) = "EC NR " Then
        sPayee = Trim$(Mid$(sPayee, 15))
        If sPayee = "" Then sPayee = v(1)
    ElseIf Left$(sPayee, 3) = "EC " Then
        sPayee = Trim$(Mid$(sPayee, 12))
        If sPayee = "" Then sPayee = v(1)
    End If
    ABNA_FindPayee = sPayee
End Function

Private Function GetSpecialPayee(sMemo As String) As String
    Dim sTmp As String
    Dim iTmp As Long

    Dim i As Integer
    i = 1
next_pattern:
    sTmp = GetMyString(IniSectionSpecialPayeeNames, _
        IniPatternPrefix & CStr(i), "")
    If sTmp = "" Then GoTo baleout
    If sMemo Like (sTmp & "*") Then
        sTmp = GetMyString(IniSectionSpecialPayeeNames, _
            IniPayeePrefix & CStr(i), "")
    Else
        i = i + 1
        GoTo next_pattern
    End If
baleout:
    GetSpecialPayee = sTmp
End Function

Public Function ABNA_FindTxnDate(t As Txn, bFound As Boolean) As Date
    Dim sMemo As String
    Dim dTmp As Date
    sMemo = t.FurtherInfo
    bFound = False
    If Left$(sMemo, 15) = "BETAALAUTOMAAT " _
    Or Left$(sMemo, 13) = "GELDAUTOMAAT " _
    Or Left$(sMemo, 9) = "CHIPKNIP " Then
        dTmp = DateSerial(CInt(Mid$(sMemo, 22, 2)) + 2000, _
            CInt(Mid$(sMemo, 19, 2)), _
            CInt(Mid$(sMemo, 16, 2)))
        dTmp = dTmp + TimeSerial(CInt(Mid$(sMemo, 25, 2)), _
            CInt(Mid$(sMemo, 28, 2)), _
            0)
    Else
        Exit Function
    End If
    bFound = True
    ABNA_FindTxnDate = dTmp
End Function

Public Function ABNA_FindServerTime(sFilein As String) As Date
    Dim iTmp As Integer
    Dim sTmp As String
    Dim dServer As Date
' extract server date/time from file name!!
    iTmp = InStrRev(sFilein, "\")
    If iTmp = 0 Then
        sTmp = sFilein
    Else
        sTmp = Mid$(sFilein, iTmp + 1)
    End If
    If Left$(sTmp, 5) = "MT940" And IsNumeric(Mid$(sTmp, 6, 12)) Then
        dServer = DateSerial(CInt(Mid$(sTmp, 10, 2)) + 2000, CInt(Mid$(sTmp, 8, 2)), CInt(Mid$(sTmp, 6, 2)))
        dServer = dServer + TimeSerial(CInt(Mid$(sTmp, 12, 2)), CInt(Mid$(sTmp, 14, 2)), CInt(Mid$(sTmp, 16, 2)))
    Else
        Debug.Print "Cannot derive server timestamp from filename"
        dServer = Now
    End If
    ABNA_FindServerTime = dServer
End Function

