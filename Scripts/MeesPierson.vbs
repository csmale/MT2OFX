Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/MeesPierson.vbs 1     5/03/06 22:56 Colin $"
Private Const FormatName = "MeesPierson Bank Nederland MT940 format"

Sub Initialise()
    LogProgress Bcfg.IDString, "Initialise"
	If Not CheckVersion() Then
		Abort
	End If
End Sub

Function DescriptiveName()
	DescriptiveName = FormatName
End Function

Sub StartSession()
    LogProgress Bcfg.IDString, "StartSession"
End Sub

Sub ProcessStatement(s)
    LogProgress Bcfg.IDString, "ProcessStatement"
End Sub

Sub ProcessTransaction(t)
	Dim dTxn
	Dim bFound

    LogProgress Bcfg.IDString, "ProcessTransaction"
	t.Payee = MPNL_FindPayee(t)
	If Cfg.ScriptDebugLevel > 5 Then
		MsgBox "Found payee: " & t.Payee
	End If
	dTxn = MPNL_FindTxnDate(t, bFound)
	If bFound then
		If Cfg.ScriptDebugLevel > 5 Then
			MsgBox "Found Txn Date: " & dTxn
		End If
		t.TxnDate = dTxn
	End If
End Sub

Sub EndSession()
    LogProgress Bcfg.IDString, "EndSession"
	If Cfg.TxnDumpFile <> "" Then
		DumpObjects Cfg.TxnDumpFile
	End If
End Sub

Public Function MPNL_FindPayee(t)
    Dim v
    Dim sPayee
    Dim iTmp
    Dim sMemo

    sPayee = t.Str86.Payee
    If sPayee = "" Then sPayee = t.Str86.GetField(sfDetails0 + 2)
    If sPayee = "" Then sPayee = "?"
    MPNL_FindPayee = sPayee
End Function

Public Function MPNL_FindTxnDate(t, bFound)
    Dim sTmp
    Dim dTmp
    Dim sMemo
    Dim re
    bFound = False
    
    sTmp = t.Str86.GetField(sfDetails0 + 1)
    If Left(sTmp, 6) <> "OPNAME" And Left(sTmp, 7) <> "BETAALD" Then Exit Function
    sMemo = sTmp
    ' Example: BETAALD  11-01-03 14U58 343R03
    
    Set re = New RegExp
    Dim Matches
    re.Pattern = "^.* (\d\d)-(\d\d)-(\d\d) (\d\d)[Uu:](\d\d).*"
    Set Matches = re.Execute(sMemo)
    If Matches.Count = 0 Then
        Debug.Print "Unable to find txn date in " & sMemo
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
    MPNL_FindTxnDate = dTmp
End Function
