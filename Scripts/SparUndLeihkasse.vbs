' MT2OFX Processing Script for Spar & Leihkasse Bank, Thayngen, Switzerland (RBABCH22866)

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/SparUndLeihkasse.vbs 2     14/02/05 22:31 Colin $"
Private Const FormatName = "Spar & Leihkasse Bank, Thayngen, Switzerland MT940 format"

Sub ConcatMemo(t, s)
	If s = "" Then
		Exit Sub
	End If
	If Len(t.Memo) > 0 Then
		t.Memo = t.Memo & Cfg.MemoDelimiter
	End If
	t.Memo = t.Memo & s
End Sub

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

	RBAB_FindDescription t
	t.Payee = RBAB_FindPayee(t)
	If Cfg.ScriptDebugLevel > 5 Then
		MsgBox "Found payee: " & t.Payee
	End If
	dTxn = RBAB_FindTxnDate(t, bFound)
	If bFound then
		If Cfg.ScriptDebugLevel > 5 Then
			MsgBox "Found Txn Date: " & dTxn
		End If
		t.TxnDate = dTxn
		t.TxnDateValid = True
	End If
	Select Case t.Str86.GetField(sfBookingText)
	Case "BANCOMAT/MAESTR"
		t.TxnType = "POS"
	Case "SALAER/RENTE"
		t.TxnType = "INT"
	Case "GELDAUTOMAT"
		t.TxnType = "ATM"
		t.Payee = "Geldautomat"
	Case "SPESEN/GEBUEHR"
		t.TxnType = "SRVCHG"
	End Select
End Sub

Sub EndSession()
    LogProgress Bcfg.IDString, "EndSession"
	If Cfg.TxnDumpFile <> "" Then
		DumpObjects Cfg.TxnDumpFile
	End If
End Sub

Public Sub RBAB_FindDescription(t)
	Dim sTmp
	t.Memo = ""
	sTmp = Trim(t.Str86.GetField(sfBookingText))
	ConcatMemo t, sTmp
    sTmp = Trim(t.Str86.GetField(sfDetails0))
    ConcatMemo t, sTmp
    sTmp = Trim(t.Str86.GetField(sfDetails0+1))
    ConcatMemo t, sTmp
    sTmp = Trim(t.Str86.GetField(sfDetails0+2))
    ConcatMemo t, sTmp
End Sub

Public Function RBAB_FindPayee(t)
	Dim sPayee
    sPayee = Trim(t.Str86.GetField(sfDetails0))
    If Len(sPayee) = 0 Then
    	sPayee = Trim(t.Str86.GetField(sfBookingText))
    End If
    RBAB_FindPayee = sPayee
End Function

Public Function RBAB_FindTxnDate(t, bFound)
    Dim sTmp
    Dim dTmp
    Dim sDate
    Dim re
    bFound = False
        
    sDate = Trim(t.Str86.GetField(21))
    If Len(sDate) = 0 Then
    	Exit Function
    End If
    Set re = New RegExp
    Dim Matches
    re.Pattern = "^(\d\d).(\d\d).(\d\d) ?/ ?(\d\d):(\d\d).*"
    Set Matches = re.Execute(sDate)
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
    RBAB_FindTxnDate = dTmp
End Function
