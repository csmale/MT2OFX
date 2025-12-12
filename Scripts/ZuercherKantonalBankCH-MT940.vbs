' MT2OFX Processing Script for Zuercher Kantonalbank (ZKBKCHZZ)

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/ZuercherKantonalBankCH-MT940.vbs 2     7/12/06 15:07 Colin $"
Private Const FormatName = "Zuercher Kantonalbank MT940 format"

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
' account number field (:25:) starts with two spaces, then 3-digit account type, another 3-digit code (possibly
' the branch) and the account number.
	s.Acct = Mid(Trim(s.Acct), 7)
	
' ZKBK leaves out the currency in the closing balance
' According to MT940 this is a mandatory field!!
    If s.ClosingBalance.Ccy = "" Then
    	s.ClosingBalance.Ccy = s.OpeningBalance.Ccy
    End If
End Sub

Sub ProcessTransaction(t)
	LogProgress Bcfg.IDString, "ProcessTransaction"

	t.Payee = ZKBK_FindPayee(t)
	If Cfg.ScriptDebugLevel > 5 Then
		MsgBox "Found payee: " & t.Payee
	End If
	If t.Amt < 0 Then
		t.TxnType = "PAYMENT"
	Else
		t.TxnType = "DEP"
	End If
End Sub

Sub EndSession()
    LogProgress Bcfg.IDString, "EndSession"
	If Cfg.TxnDumpFile <> "" Then
		DumpObjects Cfg.TxnDumpFile
	End If
End Sub

Public Function ZKBK_FindPayee(t)
    Dim v
    Dim sPayee
    Dim iTmp
    Dim sMemo

	ZKBK_FindPayee = t.NonSwift.GetField(17)
End Function
