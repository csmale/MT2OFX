Attribute VB_Name = "BankSNSB"
Option Explicit

' $Header: /MT2OFX/BankSNSB.bas 2     4/04/04 21:24 Colin $

Public Function SNSB_FindPayee(t As Txn) As String
    Dim v As Variant
    Dim sPayee As String
    Dim iTmp As Integer
    Dim sMemo As String
    Dim re As RegExp
    Dim sPat As String

' we will use this opportunity to set up a FITID for SNS as there are
' no reliable values in the MT940 to use as a basis!
    If Left$(t.Statement.OFXStatementID, 1) <> "{" Then
        t.Statement.OFXStatementID = GenerateGUID()
    End If

    sMemo = t.FurtherInfo
    If sMemo = "" Then
        SNSB_FindPayee = sMemo
        Exit Function
    End If
    v = Split(sMemo, Cfg.MemoDelimiter)
    If Not IsArray(v) Then
        SNSB_FindPayee = sMemo
        Exit Function
    End If
    sPayee = v(0)
    If IsNumeric(Left$(sPayee, 1)) And InStr(sPayee, " ") > 8 Then ' starts with account number
        sPayee = Trim$(Mid$(sPayee, 11))
    Else
        Set re = New RegExp
        Dim Matches As MatchCollection
        re.Pattern = "^(.*) [ \d]\d\.\d\d.*"
        Set Matches = re.Execute(sPayee)
        If Matches.Count > 0 Then
            If Matches(0).SubMatches.Count > 0 Then
                sPayee = Matches(0).SubMatches(0)
            End If
        End If
    End If
    SNSB_FindPayee = sPayee
    t.Reference = GenerateGUID()
End Function

Public Function SNSB_FindTxnDate(t As Txn, bFound As Boolean) As Date
    Dim dTmp As Date
    Dim sMemo As String
    Dim re As RegExp
    sMemo = t.FurtherInfo
    bFound = False
    
    If t.BookingCode <> "NBEA" And t.BookingCode <> "NGEA" Then Exit Function
    
    Set re = New RegExp
    Dim Matches As MatchCollection
    re.Pattern = "^.* ([ \d]? ?\d) ?\.( ?\d ?\d) ?\.( ?\d ?\d ?\d ?\d)  ?([ \d]? ?\d) ?u( ?\d ?\d).*"
    Set Matches = re.Execute(sMemo)
    If Matches.Count = 0 Then
        Debug.Print "Unable to find txn date in " & sMemo
        Exit Function
    End If
    Dim iDay As Integer, iMon As Integer, iYear As Integer
    Dim iHour As Integer, iMin As Integer
    iDay = CInt(Replace(Matches(0).SubMatches(0), " ", ""))
    iMon = CInt(Replace(Matches(0).SubMatches(1), " ", ""))
    iYear = CInt(Replace(Matches(0).SubMatches(2), " ", ""))
    iHour = CInt(Replace(Matches(0).SubMatches(3), " ", ""))
    iMin = CInt(Replace(Matches(0).SubMatches(4), " ", ""))
    
    dTmp = DateSerial(iYear, iMon, iDay) + TimeSerial(iHour, iMin, 0)
    
    bFound = True
    SNSB_FindTxnDate = dTmp
End Function


