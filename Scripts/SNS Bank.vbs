Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/SNS Bank.vbs 10    30/11/10 23:13 Colin $"
Private Const FormatName = "SNS Bank Nederland MT940 format"

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

' we will use this opportunity to set up a FITID for SNS as there are
' no reliable values in the MT940 to use as a basis!
'        s.OFXStatementID = MakeGUID()
        s.BankName = "SNSBNL2A"
End Sub

Sub ProcessTransaction(t)
	Dim dTxn
	Dim bFound

    LogProgress Bcfg.IDString, "ProcessTransaction"

	t.Payee = SNSB_FindPayee(t)
	If Cfg.ScriptDebugLevel > 5 Then
		MsgBox "Found payee: " & t.Payee
	End If
	dTxn = SNSB_FindTxnDate(t, bFound)
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

Public Function SNSB_FindPayee(t)
    Dim v
    Dim sPayee
    Dim iTmp
    Dim sMemo
    Dim re
    Dim sPat
    
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
    If IsNumeric(Left(sPayee, 1)) And InStr(sPayee, " ") > 8 Then ' starts with account number
        sPayee = Trim(Mid(sPayee, 11))
    Else
        If t.BookingCode = "NBEA" Then
            sPayee = Trim(Left(v(2), 23))
        Else
        If Len(sPayee) = 0 Then sPayee = v(2)
        Set re = New RegExp
        Dim Matches
        re.Pattern = "^(.*) [ \d]\d\.\d\d.*"
'msgbox "finding payee with '" & re.pattern & "' in '" & sPayee & "'"
        Set Matches = re.Execute(sPayee)
        If Matches.Count > 0 Then
            If Matches(0).SubMatches.Count > 0 Then
                sPayee = Matches(0).SubMatches(0)
'            Else
'msgbox "unable to find payee with '" & re.pattern & "' in '" & sPayee & "'"
            End If
        End If
        End If
    End If
    SNSB_FindPayee = sPayee
End Function

Public Function SNSB_FindTxnDate(t, bFound)
    Dim dTmp
    Dim sMemo
    Dim re
    sMemo = t.FurtherInfo
    bFound = False
    
    If t.BookingCode <> "NBEA" And t.BookingCode <> "NGEA" And t.BookingCode <> "NCHP" Then Exit Function
    
    Set re = New RegExp
    Dim Matches
    re.Pattern = "^.* ([ \d]? ?\d) ?\.( ?\d ?\d) ?\.( ?\d? ?\d? ?\d ?\d)  ?([ \d]? ?\d) ?[Uu]( ?\d ?\d).*"
    Set Matches = re.Execute(sMemo)
    If Matches.Count = 0 Then
'0653063431 de kappery c.v.;Arnhemseweg 146 2 7335DN APELDOORN;kaps. j. nieuwenhuis 9969209 094.011718 04.04.03 18u23 kv005;3629010967300199900362901
' SUPER DE BOER 7766     VELP GLD   4.01.2010 15U16 KV005 0WST14   
'        Message True, True, "Unable to find txn date in " & sMemo, "Check Script"
        Exit Function
    End If
    Dim iDay, iMon, iYear
    Dim iHour, iMin
    iDay = CInt(Replace(Matches(0).SubMatches(0), " ", ""))
    iMon = CInt(Replace(Matches(0).SubMatches(1), " ", ""))
    iYear = CInt(Replace(Matches(0).SubMatches(2), " ", ""))
    If iYear < 100 Then
    	If iYear > 50 Then
    		iYear = iYear + 1900
    	Else
    		iYear = iYear + 2000
    	End If
    End If
    iHour = CInt(Replace(Matches(0).SubMatches(3), " ", ""))
    iMin = CInt(Replace(Matches(0).SubMatches(4), " ", ""))
    
    dTmp = DateSerial(iYear, iMon, iDay) + TimeSerial(iHour, iMin, 0)
    
    bFound = True
    SNSB_FindTxnDate = dTmp
End Function
