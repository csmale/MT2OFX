Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/Credit Suisse.vbs 2     14/02/05 22:31 Colin $"
Private Const FormatName = "Credit Suisse Bank MT940 format"

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
	Dim dTxn
	Dim bFound
	LogProgress Bcfg.IDString, "ProcessTransaction"
	t.Payee = FindPayee(t)
	t.FITID = t.BankReference
End Sub

Sub EndSession()
	LogProgress Bcfg.IDString, "EndSession"
	If Cfg.TxnDumpFile <> "" Then
		DumpObjects Cfg.TxnDumpFile
	End If
End Sub

Function TrimLeadingDigits(s)
	Dim r
	Set r=New regexp
	r.Global = False
	r.Pattern = "^\d+ *(.*?)$"
	Dim m
	Set m=r.Execute(s)
	If m.Count = 0 Then
		TrimLeadingDigits = s
	Else
		TrimLeadingDigits = m(0).SubMatches(0)
	End If
End Function

Function FindPayee(t)
    Dim v
    Dim sPayee
    Dim iTmp
    Dim sMemo
    
    sMemo = t.Memo
    If sMemo = "" Then
        FindPayee = sMemo
        Exit Function
    End If
    v = Split(sMemo, Cfg.MemoDelimiter)
    If IsArray(v) Then
    	sPayee = v(0)
    Else
        sPayee = sMemo
        Exit Function
    End If
    iTmp = InStr(sPayee, "//")
    If iTmp > 0 Then
    	sPayee = Left(sPayee, iTmp-1)
    End If
    sPayee = TrimLeadingDigits(sPayee)
    FindPayee = sPayee
End Function
