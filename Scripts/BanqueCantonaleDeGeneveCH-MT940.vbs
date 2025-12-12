' MT2OFX Processing Script for Banque Cantonale de Genève (BCGECHGG)

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/BanqueCantonaleDeGeneveCH-MT940.vbs 1     24/11/09 22:12 Colin $"
Private Const FormatName = "Banque Cantonale de Genève MT940 format"

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
    s.BankName = "BCGECHGG"
End Sub

Sub ProcessTransaction(t)
	LogProgress Bcfg.IDString, "ProcessTransaction"

    t.Payee = BCGE_FindPayee(t)
    If Cfg.ScriptDebugLevel > 5 Then
        MsgBox "Found payee: " & t.Payee
    End If
    If t.TxnType = "" Then
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
    If t.FITID = "NONREF" Then t.FITID = ""
End Sub

Sub EndSession()
    LogProgress Bcfg.IDString, "EndSession"
    If Cfg.TxnDumpFile <> "" Then
        DumpObjects Cfg.TxnDumpFile
    End If
End Sub

Public Function BCGE_FindPayee(t)
    Dim sPayee
    Dim sMemo, sTmp, sTxnDate
    Dim iStart, iEnd
    sPayee = ""
    sMemo = t.Memo
' Achat Maestro 21.10.2009 06:56 GARE CFF GENEVE AERO Numero de car
' Bancomat 06.10.2009 17:04 CS GE MEYRIN MEDIA Numero de carte: 784
    iStart = InStr(sMemo, "Achat Maestro")
    If iStart = 0 Then
        iStart = InStr(sMemo, "Bancomat")
        If iStart > 0 Then
            t.TxnType = "ATM"
            sTmp = Mid(sMemo, 9) ' lose Bancomat
        End If
    Else
        t.TxnType = "POS"
        sTmp = Mid(sMemo, 15) ' lose Achat Maestro
    End If
    If iStart > 0 Then
        sTxnDate = Left(sTmp, 16)
        t.TxnDate = ParseTxnDate(sTxnDate)
        t.TxnDateValid = (t.TxnDate <> NODATE)
        sPayee = Mid(sTmp, 18) ' lose transaction date
        iEnd = InStr(sPayee, "Numero de ")
        If iEnd > 0 Then
            sPayee = Trim(Left(sPayee, iEnd-1))
        End If
    End If
   
	BCGE_FindPayee = sPayee
End Function

Function ParseTxnDate(sDate)
    Dim dTmp
    Dim sTmp
    dTmp = ParseDateEx(Left(sDate, 10), "DMY", ".")
    If dTmp <> NODATE Then
        dTmp = dTmp + TimeSerial(CInt(Mid(sDate, 12, 2)), CInt(Mid(sDate, 15, 2)), 0)
' msgbox "date: " & sDate & "-->" & dTmp
    End If
    ParseTxnDate = dTmp
End Function
