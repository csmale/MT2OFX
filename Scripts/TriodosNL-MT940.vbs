' MT2OFX Processing Script for Triodos Bank, Netherlands (TRIONL21XXX)

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/TriodosNL-MT940.vbs 1     12/01/06 21:33 Colin $"
Private Const FormatName = "Triodos Bank, Netherlands MT940 format"

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
' file does not contain a proper bank identifier! it just has ":21:TRIODOSBANK/..."
    s.BankName = "TRIONL21"
End Sub

Sub ProcessTransaction(t)
	Dim dTxn
	Dim bFound

	LogProgress Bcfg.IDString, "ProcessTransaction"

	TRIO_FindDescription t
	t.Payee = TRIO_FindPayee(t)
	If Cfg.ScriptDebugLevel > 5 Then
		MsgBox "Found payee: " & t.Payee
	End If
	dTxn = TRIO_FindTxnDate(t, bFound)
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

Public Sub TRIO_FindDescription(t)
	Dim sTmp
	t.Memo = ""
	sTmp = Trim(t.Str86.GetField(sfBookingText))
    sTmp = sTmp & Trim(t.Str86.GetField(sfDetails0))
    sTmp = sTmp & Trim(t.Str86.GetField(sfDetails0+1))
    sTmp = sTmp & Trim(t.Str86.GetField(sfDetails0+2))
    sTmp = sTmp & Trim(t.Str86.GetField(sfDetails0+3))
    sTmp = sTmp & Trim(t.Str86.GetField(sfDetails0+4))
	t.Memo = sTmp
	ConcatMemo Trim(t.Str86.GetField(sfAcctPayee))
End Sub

Public Function TRIO_FindPayee(t)
	Dim sPayee
    sPayee = Trim(t.Str86.GetField(sfDetails0))
    If Len(sPayee) = 0 Then
    	sPayee = Trim(t.Str86.GetField(sfBookingText))
    End If
    TRIO_FindPayee = sPayee
End Function

Public Function TRIO_FindTxnDate(t, bFound)
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
    TRIO_FindTxnDate = dTmp
End Function
