' MT2OFX Processing Script for St.Galler Kantonalbank (SGKBCH22)

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/StGallerKantonalbankCH-MT940.vbs 1     11/03/08 9:01 Colin $"
Private Const FormatName = "St.Galler Kantonalbank MT940 format"

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
	s.BankName = "SGKBCH22"
	s.Acct = Trim(s.Acct)
End Sub

Sub ProcessTransaction(t)
	Dim dTxn
	Dim bFound

	LogProgress Bcfg.IDString, "ProcessTransaction"

	t.Memo = SGKB_FindDescription(t)
	t.Payee = SGKB_FindPayee(t)
	If Cfg.ScriptDebugLevel > 5 Then
		MsgBox "Found payee: " & t.Payee
	End If
	dTxn = SGKB_FindTxnDate(t, bFound)
	If bFound Then
		If Cfg.ScriptDebugLevel > 5 Then
			MsgBox "Found Txn Date: " & dTxn
		End If
		t.TxnDate = dTxn
		t.TxnDateValid = True
	End If
	If t.TxnType = "" Then
		If t.Amt > 0 Then
			t.TxnType = "DEP"
		Else
			t.TxnType = "PAYMENT"
		End If
	End If
End Sub

Sub EndSession()
    LogProgress Bcfg.IDString, "EndSession"
	If Cfg.TxnDumpFile <> "" Then
		DumpObjects Cfg.TxnDumpFile
	End If
End Sub

Public Function SGKB_FindDescription(t)
	Dim i
	Dim sTmp, sDesc
	sDesc = ""
	For i=nsfApplication0 To nsfApplicationLast
' Don't know why we need the CInt() - but it avoids a Type Mismatch error!
		sTmp = t.NonSwift.GetField(CInt(i))
		If (sTmp <> "") Or (Not Bcfg.SkipEmptyMemoFields) Then
			If sDesc = "" Then
				sDesc = sTmp
			Else
				sDesc = sDesc & Cfg.MemoDelimiter & sTmp
			End If
		End If
	Next
	SGKB_FindDescription = sDesc
End Function

Public Function SGKB_FindPayee(t)
    Dim v
    Dim sPayee
    Dim iTmp
    Dim sMemo

    Select Case t.NonSwift.GetField(nsfBookingText)
    Case "Maestro-Bezug"
	    sPayee = t.NonSwift.GetField(1)
	    t.TxnType = "POS"
	Case "BM-Bezug CHF", "BM-Bezug EUR"
		sPayee = "Cash Withdrawal"
		t.TxnType = "ATM"
	Case "Postgiro"
		sPayee = t.NonSwift.GetField(2)
	Case "Belastung LSV+"
		sPayee = t.NonSwift.GetField(2)
	Case "Vergütung"	
		sPayee = t.NonSwift.GetField(1)
		t.TxnType = "CREDIT"
	Case "Gebühr BM Fremdw"
		t.TxnType = "FEE"
	Case Else
		sPayee = ""
    End Select
    SGKB_FindPayee = sPayee
End Function

Public Function SGKB_FindTxnDate(t, bFound)
    Dim sTmp
    Dim dTmp
    Dim sMemo
    Dim re
    bFound = False
        
    sMemo = t.Memo
    
    Set re = New RegExp
    Dim Matches
' NS info:
' 02Karten Nr. 78222278 Zeit: 10.11
' 03Datum: 04.01.2008
	sTmp = t.NonSwift.GetField(2) & t.NonSwift.GetField(3)
    re.Pattern = "^Karten.*Zeit: (\d\d).(\d\d).*Datum: (\d\d).(\d\d).(\d\d\d\d).*"
    Set Matches = re.Execute(sTmp)
    If Matches.Count = 0 Then
'        MsgBox "Unable to find txn date in " & sTmp
        Exit Function
    End If

    Dim iDay, iMon, iYear
    Dim iHour, iMin
    iHour = CInt(Matches(0).SubMatches(0))
    iMin = CInt(Matches(0).SubMatches(1))
    iDay = CInt(Matches(0).SubMatches(2))
    iMon = CInt(Matches(0).SubMatches(3))
    iYear = CInt(Matches(0).SubMatches(4))
    
    dTmp = DateSerial(iYear, iMon, iDay) + TimeSerial(iHour, iMin, 0)
    
    bFound = True
    SGKB_FindTxnDate = dTmp
End Function
