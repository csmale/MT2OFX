Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/Van Lanschot.vbs 5     14/02/05 22:31 Colin $"
Private Const FormatName = "Van Lanschot Bank MT940 format"

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
    If sPayee = "" Then
    	sPayee = "?"
    End If
    t.Payee = sPayee
End Sub

Sub EndSession()
    LogProgress Bcfg.IDString, "EndSession"
	If Cfg.TxnDumpFile <> "" Then
		DumpObjects Cfg.TxnDumpFile
	End If
End Sub
