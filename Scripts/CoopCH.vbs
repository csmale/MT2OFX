' MT2OFX Processing Script for Coop Bank Switzerland (COOPCHBB)

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/CoopCH.vbs 2     14/02/05 22:31 Colin $"
Private Const FormatName = "Coop Bank Switzerland MT940 format"

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
End Sub

Sub ProcessTransaction(t)
	Dim dTxn
	Dim bFound

	LogProgress Bcfg.IDString, "ProcessTransaction"

	t.FurtherInfo = COOP_FindDescription(t)
	t.Payee = COOP_FindPayee(t)
	If Cfg.ScriptDebugLevel > 5 Then
		MsgBox "Found payee: " & t.Payee
	End If
	dTxn = COOP_FindTxnDate(t, bFound)
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

Public Function COOP_FindDescription(t)
	Dim sDesc
	Dim sTmp
	sDesc = t.Str86.GetField(sfNamePayee)
    sTmp = Trim(t.Str86.GetField(sfNamePayee2))
    If Len(sTmp) > 0 Then
    	sDesc = sDesc & Cfg.MemoDelimiter & sTmp
    End If
    If Len(sDesc) = 0 Then
    	sDesc = Trim(t.Str86.GetField(sfBookingText))
    End If
	COOP_FindDescription = sDesc
End Function

Public Function COOP_FindPayee(t)
	Dim sPayee
    sPayee = Trim(t.Str86.GetField(sfNamePayee))
    If Len(sPayee) = 0 Then
    	sPayee = Trim(t.Str86.GetField(sfBookingText))
    End If
    If Len(sPayee) = 0 Then
    	sPayee = "Unbekannt"
    End If
    COOP_FindPayee = sPayee
End Function

Public Function COOP_FindTxnDate(t, bFound)
    Dim sTmp
    Dim dTmp
    Dim sMemo
    Dim re
    bFound = False

' Not enough info to do this for COOP - this is an example for another bank!
    Exit Function
        
    sMemo = t.FurtherInfo
    
    Set re = New RegExp
    Dim Matches
    re.Pattern = "^.*(\d\d).(\d\d).(\d\d).*"
    Set Matches = re.Execute(sMemo)
    If Matches.Count = 0 Then
'        Debug.Print "Unable to find txn date in " & sMemo
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
    COOP_FindTxnDate = dTmp
End Function
