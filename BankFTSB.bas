Attribute VB_Name = "BankFTSB"
Option Explicit

' $Header: /MT2OFX/BankFTSB.bas 2     4/04/04 21:24 Colin $

Public Function FTSB_FindPayee(t As Txn) As String
    Dim v As Variant
    Dim sPayee As String
    Dim iTmp As Integer
    Dim sMemo As String

    sPayee = t.Str86.Payee
    If sPayee = "" Then sPayee = t.Str86.GetField(sfDetails0 + 2)
    If sPayee = "" Then sPayee = "?"
    FTSB_FindPayee = sPayee
End Function

Public Function FTSB_FindTxnDate(t As Txn, bFound As Boolean) As Date
    Dim sTmp As String
    Dim dTmp As Date
    Dim sMemo As String
    Dim re As RegExp
    bFound = False
    
    sTmp = t.Str86.GetField(sfDetails0 + 1)
    If Left$(sTmp, 6) <> "OPNAME" And Left$(sTmp, 7) <> "BETAALD" Then Exit Function
    sMemo = sTmp
    ' BETAALD  11-01-03 14U58 343R03
    
    Set re = New RegExp
    Dim Matches As MatchCollection
    re.Pattern = "^.* (\d\d)-(\d\d)-(\d\d) (\d\d)[Uu:](\d\d).*"
    Set Matches = re.Execute(sMemo)
    If Matches.Count = 0 Then
        Debug.Print "Unable to find txn date in " & sMemo
        Exit Function
    End If
    Dim iDay As Integer, iMon As Integer, iYear As Integer
    Dim iHour As Integer, iMin As Integer
    iDay = CInt(Matches(0).SubMatches(0))
    iMon = CInt(Matches(0).SubMatches(1))
    iYear = CInt(Matches(0).SubMatches(2))
    If iYear < 50 Then
        iYear = iYear + 2000
    Else
        iYear = iYear + 1900
    End If
    iHour = CInt(Matches(0).SubMatches(3))
    iMin = CInt(Matches(0).SubMatches(4))
    
    dTmp = DateSerial(iYear, iMon, iDay) + TimeSerial(iHour, iMin, 0)
    
    bFound = True
    FTSB_FindTxnDate = dTmp
End Function

