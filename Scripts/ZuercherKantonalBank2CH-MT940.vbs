' MT2OFX Processing Script for Zuercher Kantonalbank (ZKBKCHZZ)

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/ZuercherKantonalBank2CH-MT940.vbs 3     11/07/09 23:05 Colin $"
Private Const FormatName = "Zuercher Kantonalbank MT940 format as from 25 March 2008"

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
	s.BankName = "ZKBKCHZZ"
' account number field (:25:) starts with a 3-digit account type, another 3-digit code (possibly
' the branch) and the account number.
	s.Acct = Mid(Trim(s.Acct), 8)
End Sub

Sub ProcessTransaction(t)
	LogProgress Bcfg.IDString, "ProcessTransaction"

	t.Payee = ZKBK_FindPayee(t)
	If Cfg.ScriptDebugLevel > 5 Then
		MsgBox "Found payee: " & t.Payee
	End If
	If t.BookingCode = "NCHG" Or t.BookingCode = "NTRF" Then
		If t.Amt < 0 Then
			t.TxnType = "PAYMENT"
		Else
			t.TxnType = "DEP"
		End If
	End If
    If Len(t.SupplementaryDetails) > 0 Then
        t.Memo = t.SupplementaryDetails & Cfg.MemoDelimiter & t.Memo
    End If
	t.FITID = t.BankReference
End Sub

Sub EndSession()
    LogProgress Bcfg.IDString, "EndSession"
	If Cfg.TxnDumpFile <> "" Then
		DumpObjects Cfg.TxnDumpFile
	End If
End Sub

Public Function ZKBK_FindPayee(t)
    Dim sPayee
	Dim sMemo
	Dim iStart, iEnd
	sPayee = ""
	sMemo = t.Memo
	If t.BookingCode = "NDDT" Then
		iStart = 14
		iEnd = InStr(iStart+1, sMemo, Cfg.MemoDelimiter)
		If iEnd > 0 Then
			sPayee = Mid(sMemo, iStart+1, iEnd - iStart - 1)
		End If
	Else
		iStart = InStr(sMemo, Cfg.MemoDelimiter)
		If iStart > 0 Then
			iEnd = InStr(iStart+1, sMemo, Cfg.MemoDelimiter)
			If iEnd > 0 Then
				sPayee = Mid(sMemo, iStart+1, iEnd - iStart - 1)
			End If
		End If
	End If
	ZKBK_FindPayee = sPayee
End Function
