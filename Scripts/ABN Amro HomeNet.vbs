Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/ABN Amro HomeNet.vbs 3     4/04/04 21:42 Colin $"

Const IniSectionSpecialPayeeNames = "SpecialPayeeNames"
Const IniPatternPrefix = "Pattern"
Const IniPayeePrefix = "Payee"

DIM NLG2EUR

Sub Initialise()
	LogProgress Bcfg.IDString, "Initialise"
	If Not CheckVersion() Then
		Abort()
	End If
End Sub

Sub StartSession()
	LogProgress Bcfg.IDString, "StartSession"
End Sub

Sub ProcessStatement(s)
	LogProgress Bcfg.IDString, "ProcessStatement"
	If s.ClosingBalance.Ccy = "NLG" Then
		s.ClosingBalance.Ccy = "EUR"
		s.ClosingBalance.Amt = s.ClosingBalance.Amt / 2.20371
		NLG2EUR = True
	Else
		NLG2EUR = False
	End If
	If s.OpeningBalance.Ccy = "NLG" Then
		s.OpeningBalance.Ccy = "EUR"
		s.OpeningBalance.Amt = s.OpeningBalance.Amt / 2.20371
	End If
	If s.AvailableBalance.Ccy = "NLG" Then
		s.AvailableBalance.Ccy = "EUR"
		s.AvailableBalance.Amt = s.AvailableBalance.Amt / 2.20371
	End If
End Sub

Sub ProcessTransaction(t)
	Dim dTxn
	Dim bFound
	LogProgress Bcfg.IDString, "ProcessTransaction"
	t.Payee = ABNA_FindPayee(t)
	If Cfg.ScriptDebugLevel > 5 Then
		MsgBox "Found payee: " & t.Payee
	End If
	dTxn = ABNA_FindTxnDate(t, bFound)
	If bFound then
		If Cfg.ScriptDebugLevel > 5 Then
			MsgBox "Found Txn Date: " & dTxn
		End If
		t.TxnDate = dTxn
	End If
	If NLG2EUR Then
		t.Amt = t.Amt / 2.20371
	End If
End Sub

Sub EndSession()
	LogProgress Bcfg.IDString, "EndSession"
	If Cfg.TxnDumpFile <> "" Then
		DumpObjects Cfg.TxnDumpFile
	End If
End Sub

Function ABNA_FindPayee(t)
    Dim v
    Dim sPayee
    Dim iTmp
    Dim sMemo
    
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
    If Left(sPayee, 16) = "PROV.  TELEGIRO " Then
        sPayee = Mid(sPayee, 17)
    End If
    If Left(sPayee, 14) = "BETAALAUTOMAAT" or Left(sPayee, 4) = "BEA " Then
        sPayee = v(1)
        iTmp = InStr(sPayee, ",")
        If iTmp <> 0 Then
            sPayee = Trim(Left(sPayee, iTmp - 1))
        End If
    ElseIf IsNumeric(Left(sPayee, 1)) Then ' starts with account number
        sPayee = Trim(Mid(sPayee, 14))
        If sPayee = "" Then sPayee = v(1)
    ElseIf Left(sPayee, 5) = "GIRO " Then
        sPayee = Mid(sPayee, 6)
        While Len(sPayee) > 0 And (Left(sPayee, 1) = " " Or IsNumeric(Left(sPayee, 1)))
            sPayee = Mid(sPayee, 2)
        Wend
        sPayee = Trim(sPayee)
        If sPayee = "" Then
            sPayee = v(1)
        End If
    ElseIf Left(sPayee, 2) = "NI" And IsNumeric(Mid(sPayee, 3, 1)) Then
        sPayee = v(1)
    ElseIf Left(sPayee, 6) = "EC NR " Then
        sPayee = Trim(Mid(sPayee, 15))
        If sPayee = "" Then sPayee = v(1)
    ElseIf Left(sPayee, 3) = "EC " Then
        sPayee = Trim(Mid(sPayee, 12))
        If sPayee = "" Then sPayee = v(1)
    End If
    ABNA_FindPayee = sPayee
End Function

Function ABNA_FindTxnDate(t, bFound)
    Dim sMemo
    Dim dTmp
    sMemo = t.FurtherInfo
    bFound = False
    If Left(sMemo, 15) = "BETAALAUTOMAAT " _
    Or Left(sMemo, 13) = "GELDAUTOMAAT " _
    Or Left(sMemo, 9) = "CHIPKNIP " Then
        dTmp = DateSerial(CInt(Mid(sMemo, 22, 2)) + 2000, _
            CInt(Mid(sMemo, 19, 2)), _
            CInt(Mid(sMemo, 16, 2)))
        dTmp = dTmp + TimeSerial(CInt(Mid(sMemo, 25, 2)), _
            CInt(Mid(sMemo, 28, 2)), _
            0)
    Elseif Left(sMemo, 4) = "BEA " _
        Or Left(sMemo, 4) = "GEA " _
        Or Left(sMemo, 5) = "CHIP " Then
            dTmp = DateSerial(CInt(Mid(sMemo, 25, 2)) + 2000, _
                CInt(Mid(sMemo, 22, 2)), _
                CInt(Mid(sMemo, 19, 2)))
            dTmp = dTmp + TimeSerial(CInt(Mid(sMemo, 28, 2)), _
                CInt(Mid(sMemo, 31, 2)), _
                0)
    Else
        Exit Function
    End If
    bFound = True
    ABNA_FindTxnDate = dTmp
End Function

Function GetSpecialPayee(sMemo)
    Dim sTmp
    Dim bDone
    Dim i
    Dim re
    Dim Matches
    
    Set re = New RegExp
    i = 1
    bDone = False
    sTmp = GetConfigString(IniSectionSpecialPayeeNames, _
        IniPatternPrefix & CStr(i), "")
    Do While sTmp <> "" And Not bDone
        re.Pattern = "^" & sTmp & ".*"
        Set Matches = re.Execute(sMemo)
        If Matches.Count > 0 Then
            sTmp = GetConfigString(IniSectionSpecialPayeeNames, _
                IniPayeePrefix & CStr(i), "")
            bDone = True
        Else
            i = i + 1
            sTmp = GetConfigString(IniSectionSpecialPayeeNames, _
                IniPatternPrefix & CStr(i), "")
        End If
    Loop
    GetSpecialPayee = sTmp
End Function
