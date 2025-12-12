Option Explicit
' MT2OFX processing script for Raiffeisenbank Zurich (RAIFCH22)

Private Const ScriptVersion = "$Header: /MT2OFX/RaiffeisenbankZurich.vbs 3     14/02/05 22:31 Colin $"
Private Const FormatName = "Raiffeisenbank Zurich MT940 format"

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
	t.Payee = RAIF_FindPayee(t)
	If Cfg.ScriptDebugLevel > 5 Then
		MsgBox "Found payee: " & t.Payee
	End If
End Sub

Sub EndSession()
	LogProgress Bcfg.IDString, "EndSession"
	If Cfg.TxnDumpFile <> "" Then
		DumpObjects Cfg.TxnDumpFile
	End If
End Sub

Function RAIF_FindPayee(t)
    Dim v
    Dim sPayee
    Dim iTmp
    Dim sMemo
    
    sMemo = t.FurtherInfo
    If sMemo = "" Then
        RAIF_FindPayee = "?"
        Exit Function
    End If
	If Mid(sMemo, 3, 1) = " " Then
		sPayee = Mid(sMemo, 4)
		Select Case Left(sMemo, 2)
		Case "BM"
			t.TxnType = "ATM"
		Case "PO"
			t.TxnType = "POS"
		Case "BE"
			t.TxnType = "POS"
		Case "GU"
			t.TxnType = "DEP"
		End Select
	Else
		sPayee = sMemo
	End If
    RAIF_FindPayee = sPayee
End Function
