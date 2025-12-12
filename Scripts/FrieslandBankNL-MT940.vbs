Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/FrieslandBankNL-MT940.vbs 1     11/10/09 15:56 Colin $"
Private Const FormatName = "Frieslandbank MT940 format"

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
End Sub

Sub ProcessTransaction(t)
    Dim sPayee
    LogProgress Bcfg.IDString, "ProcessTransaction"

    sPayee = t.Str86.Payee
    If sPayee = "" Then
    	sPayee = t.Str86.GetField(sfBookingText)
    End If
    t.Payee = Trim(sPayee)
    Dim i, sMemo
    For i = sfDetails0 To sfDetailsLast
        sMemo = sMemo & t.Str86.GetField(CInt(i))
    Next
    t.Memo = Trim(sMemo)
End Sub

Sub EndSession()
    LogProgress Bcfg.IDString, "EndSession"
	If Cfg.TxnDumpFile <> "" Then
		DumpObjects Cfg.TxnDumpFile
	End If
End Sub
