' MT2OFX Processing Script for ING Bank (INGBNL2A)

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/INGNL-MT940.vbs 1     20/02/08 0:41 Colin $"
Private Const FormatName = "ING Bank Nederland MT940 format"

Sub Initialise()
    LogProgress Bcfg.IDString, "Initialise"
End Sub

Function DescriptiveName()
	DescriptiveName = FormatName
End Function

Sub StartSession()
    LogProgress Bcfg.IDString, "StartSession"
End Sub

Sub ProcessStatement(s)
    LogProgress Bcfg.IDString, "ProcessStatement"
    s.StatementID = Trim(s.StatementID)
End Sub

Sub ProcessTransaction(t)
	Dim dTxn
	Dim bFound

	LogProgress Bcfg.IDString, "ProcessTransaction"

	t.Payee = ING_FindPayee(t)
	If Cfg.ScriptDebugLevel > 5 Then
		MsgBox "Found payee: " & t.Payee
	End If
	dTxn = ING_FindTxnDate(t, bFound)
	If bFound then
		If Cfg.ScriptDebugLevel > 5 Then
			MsgBox "Found Txn Date: " & dTxn
		End If
		t.TxnDate = dTxn
		t.TxnDateValid = True
	End If
	If t.BookingCode = "N030" Then
		t.TxnType = "POS"
	ElseIf t.BookingCode = "N038" Then
		t.TxnType = "DIRECTDEBIT"
	End If
End Sub

Sub EndSession()
    LogProgress Bcfg.IDString, "EndSession"
	If Cfg.TxnDumpFile <> "" Then
		DumpObjects Cfg.TxnDumpFile
	End If
End Sub

Public Function ING_FindPayee(t)
    Dim sPayee
    Dim iSep

	If IsNumeric(Left(t.Memo, 1)) Or (Left(t.Memo, 1) = "P" And IsNumeric(Mid(t.Memo, 2, 9))) Then
		sPayee = Trim(Mid(t.Memo, 11))
	ElseIf StartsWith(t.Memo, "BETAALAUTOMAAT") Then
		sPayee = Mid(t.Memo, 36)
		iSep = InStr(sPayee, ",")
		If iSep > 0 Then
			sPayee = Left(sPayee, iSep-1)
		End If
	End If
' het lijkt erop dat Postbank veel te weinig informatie in de
' bestanden stopt zodat het onderstaande de enige optie is...
	sPayee = Trim(sPayee)
    ING_FindPayee = Trim(Left(sPayee, 32))
End Function

Public Function ING_FindTxnDate(t, bFound)
    Dim sTmp
    Dim dTmp
    Dim sMemo
    Dim re
    bFound = False

    sMemo = t.Memo
    
    Set re = New RegExp
    Dim Matches
    re.Pattern = "^.*(\d\d).(\d\d).(\d\d\d\d) (\d\d):(\d\d) UUR.*"
    Set Matches = re.Execute(sMemo)
    If Matches.Count = 0 Then
'        MsgBox "Unable to find txn date in " & sMemo
        Exit Function
    End If
    Dim iDay, iMon, iYear
    Dim iHour, iMin
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
    ING_FindTxnDate = dTmp
End Function
