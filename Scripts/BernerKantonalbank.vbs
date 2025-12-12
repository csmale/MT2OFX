' MT2OFX Processing Script for Berner Kantonalbank (KBBECH22)

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/BernerKantonalbank.vbs 3     14/02/05 22:31 Colin $"
Private Const FormatName = "Berner Kantonalbank MT940 format"

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
   
' KBBE leaves out the currency in the closing balance
' According to MT940 this is a mandatory field!!
        If s.ClosingBalance.Ccy = "" Then
              s.ClosingBalance.Ccy = s.OpeningBalance.Ccy
        End If
        
' KBBE has extra spaces around the account number
	s.Acct = Trim(s.Acct)
End Sub

Sub ProcessTransaction(t)
	Dim dTxn
	Dim bFound

	LogProgress Bcfg.IDString, "ProcessTransaction"

'	t.FurtherInfo = KBBE_FindDescription(t)
	t.Payee = KBBE_FindPayee(t)
	If Cfg.ScriptDebugLevel > 5 Then
		MsgBox "Found payee: " & t.Payee
	End If
	dTxn = KBBE_FindTxnDate(t, bFound)
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

Public Function KBBE_FindDescription(t)
	Dim i
	Dim sTmp, sDesc
	sDesc = ""
	For i=nsfApplication0 To nsfApplicationLast
' Don't know why we need the CInt() - but it avoids a Type Mismatch error!
		sTmp = t.NonSwift.GetField(CInt(i))
		If sTmp = "/" Then
			sTmp = ""
		End If
		If (sTmp <> "") Or (Not Bcfg.SkipEmptyMemoFields) Then
			If sDesc = "" Then
				sDesc = sTmp
			Else
				sDesc = sDesc & Cfg.MemoDelimiter & sTmp
			End If
		End If
	Next
	KBBE_FindDescription = sDesc
End Function

Public Function KBBE_FindPayee(t)
    Dim v
    Dim sPayee
    Dim iTmp
    Dim sMemo

    sPayee = t.Str86.GetField(sfDetails0)
    If sPayee = "" Then
    	sPayee = t.Str86.GetField(sfBookingText)
    End If
    KBBE_FindPayee = sPayee
End Function

Public Function KBBE_FindTxnDate(t, bFound)
    Dim sTmp
    Dim dTmp
    Dim sMemo
    Dim re
    bFound = False
        
    sMemo = t.FurtherInfo
    
    Set re = New RegExp
    Dim Matches
    re.Pattern = "^.*(\d\d).(\d\d).(\d\d) / (\d\d):(\d\d).*"
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
    If iYear < 50 Then
        iYear = iYear + 2000
    Else
        iYear = iYear + 1900
    End If
    iHour = CInt(Matches(0).SubMatches(3))
    iMin = CInt(Matches(0).SubMatches(4))
    
    dTmp = DateSerial(iYear, iMon, iDay) + TimeSerial(iHour, iMin, 0)
    
    bFound = True
    KBBE_FindTxnDate = dTmp
End Function
