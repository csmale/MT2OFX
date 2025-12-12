' MT2OFX Processing Script for Stadtssparkasse Dortmund (DORTDE33XXX)

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/StadtsparkasseDortmund.vbs 2     14/02/05 22:31 Colin $"
Private Const FormatName = "Stadtssparkasse Dortmund MT940 format"

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

	t.Payee = DORT_FindPayee(t)
	If Cfg.ScriptDebugLevel > 5 Then
		MsgBox "Found payee: " & t.Payee
	End If
	dTxn = DORT_FindTxnDate(t, bFound)
	If bFound then
		If Cfg.ScriptDebugLevel > 5 Then
			MsgBox "Found Txn Date: " & dTxn
		End If
		t.TxnDate = dTxn
		t.TxnDateValid = True
	End If
End Sub

Sub EndSession()
    LogProgress Bcfg.IDString, "EndSession"
	If Cfg.TxnDumpFile <> "" Then
		DumpObjects Cfg.TxnDumpFile
	End If
End Sub

Public Function DORT_FindPayee(t)
    Dim v
    Dim sPayee
    Dim iTmp
    Dim sMemo

    sPayee = t.Str86.GetField(sfDetails0)
    If sPayee = "" Then
    	sPayee = t.Str86.GetField(sfBookingText)
    End If
    DORT_FindPayee = sPayee
End Function

Public Function DORT_FindTxnDate(t, bFound)
    Dim sTmp
    Dim dTmp
    Dim sMemo
    Dim re
    bFound = False
        
    sMemo = t.FurtherInfo
    
    Set re = New RegExp
    Dim Matches
    re.Pattern = "^.*DATUM (\d\d)\.(\d\d)\.(\d\d\d\d), (\d\d):(\d\d) UHR.*"
    Set Matches = re.Execute(sMemo)
    If Matches.Count = 0 Then
'        msgbox "Unable to find txn date in " & sMemo
        Exit Function
    End If
    Dim iDay, iMon, iYear
    Dim iHour, iMin
    iDay = CInt(Matches(0).SubMatches(0))
    iMon = CInt(Matches(0).SubMatches(1))
    iYear = CInt(Matches(0).SubMatches(2))
    iHour = CInt(Matches(0).SubMatches(3))
    iMin = CInt(Matches(0).SubMatches(4))
    
    dTmp = DateSerial(iYear, iMon, iDay) + TimeSerial(iHour, iMin, 0)
    
    bFound = True
    DORT_FindTxnDate = dTmp
End Function
