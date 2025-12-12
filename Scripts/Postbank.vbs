' MT2OFX Processing Script for Postbank (PSTBNL21)

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/Postbank.vbs 7     27/01/08 10:31 Colin $"
Private Const FormatName = "Postbank Nederland MT940 format"

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
' Postbank are now apparantly always using "000" in the :28C: lines instead of a proper statement number
' this causes duplicate FITID values, so zap the statement ID here and MT2OFX will create its own
' based on the statement date.
    If s.StatementID = "000" Then
    	s.StatementID = ""
    End If
End Sub

Sub ProcessTransaction(t)
	Dim dTxn
	Dim bFound

	LogProgress Bcfg.IDString, "ProcessTransaction"

' PSTB uses 3-char booking codes - non-standard!!!
	t.Reference = Right(t.BookingCode,1) & t.Reference
	t.BookingCode = Left(t.BookingCode,3)

	t.Payee = PSTB_FindPayee(t)
	If Cfg.ScriptDebugLevel > 5 Then
		MsgBox "Found payee: " & t.Payee
	End If
	dTxn = PSTB_FindTxnDate(t, bFound)
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

Public Function PSTB_FindPayee(t)
    Dim sPayee
    Dim iSep

	If IsNumeric(Left(t.Memo, 1)) Then
		sPayee = Trim(Mid(t.Memo, 12))
		If Left(sPayee, 4) = "KN: " Then
			iSep = InStr(sPayee, cfg.MemoDelimiter)
			If iSep > 0 Then
				sPayee = Mid(sPayee, iSep+1)
			End If
		Else
			iSep = InStr(sPayee, Cfg.MemoDelimiter)
			If iSep > 0 Then
				sPayee = Trim(Left(sPayee, iSep-1))
			End If
		End If
	End If
' het lijkt erop dat Postbank veel te weinig informatie in de
' bestanden stopt zodat het onderstaande de enige optie is...
    PSTB_FindPayee = Left(sPayee, 32)
End Function

Public Function PSTB_FindTxnDate(t, bFound)
    Dim sTmp
    Dim dTmp
    Dim sMemo
    Dim re
    bFound = False

' Not enough info to do this for PSTB - this is an example for another bank!
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
    PSTB_FindTxnDate = dTmp
End Function
