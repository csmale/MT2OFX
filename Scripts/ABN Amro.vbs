Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/ABN Amro.vbs 15    23/12/14 0:26 Colin $"
Private Const FormatName = "ABN AMRO MT940-formaat."
Const ScriptName = "ABNAmro"

Const IniSectionSpecialPayeeNames = "SpecialPayeeNames"
Const IniPatternPrefix = "Pattern"
Const IniPayeePrefix = "Payee"

' Added NLG conversion stuff to support imports from HomeNet and OfficeNet
' Set ConvToEur to True if you want all NLG amounts automatically converted to EUR
Dim ConvToEur
Const EURRate = 2.20371
Dim NLG2EUR
Dim QuickenBankID

' Property List is an array of arrays, each of which has the following elements:
'	1. Property key - used to reference properties
'	2. Property name - used as a label in the config screen
'	3. Property description - used as a description or tooltip in the config screen
'	4. Data type - ptString, ptBoolean, ptInteger, ptFloat, ptDate, ptChoice
'	5. Value list (will be displayed in a combobox) - array of values (Only with ptChoice)
Dim aPropertyList
aPropertyList = Array( _
	Array("ConvertNLG2EUR", "Enable conversion of NLG to EUR", _
		"Set this option to True to have all NLG amounts in the file automatically converted to euro.", _
		ptBoolean), _
	Array("QuickenBankID", "Quicken Bank ID", _
		"Bank ID to use in <INTU.BID> for Quicken", _
		ptInteger) _
	)

' function DescriptiveName
' returns a string with a descriptive name of this script
Function DescriptiveName()
	DescriptiveName = FormatName
End Function

Sub Initialise()
	LogProgress Bcfg.IDString, "Initialise"
	If Not CheckVersion() Then
		Abort
	End If
	LoadProperties ScriptName, aPropertyList
End Sub

Sub Configure
	If ShowConfigDialog(ScriptName, aPropertyList) Then
		SaveProperties ScriptName, aPropertyList
	End If
End Sub

Sub StartSession()
	LogProgress Bcfg.IDString, "StartSession"
	ConvToEur = GetProperty("ConvertNLG2EUR")
	QuickenBankID = GetProperty("QuickenBankID")
	Bcfg.IntuitBankID = QuickenBankID
	Session.ServerTime = ABNA_FindServerTime(Session.FileIn)
End Sub

Sub ProcessStatement(s)
	LogProgress Bcfg.IDString, "ProcessStatement"
	If ConvToEur Then
		If s.ClosingBalance.Ccy = "NLG" Then
			s.ClosingBalance.Ccy = "EUR"
			s.ClosingBalance.Amt = s.ClosingBalance.Amt / EURRate
			NLG2EUR = True
		Else
			NLG2EUR = False
		End If
		If NLG2EUR Then
			If s.OpeningBalance.Ccy = "NLG" Then
				s.OpeningBalance.Ccy = "EUR"
				s.OpeningBalance.Amt = s.OpeningBalance.Amt / EURRate
			End If
			If s.AvailableBalance.Ccy = "NLG" Then
				s.AvailableBalance.Ccy = "EUR"
				s.AvailableBalance.Amt = s.AvailableBalance.Amt / EURRate
			End If
		End If
	Else
		NLG2EUR = False
	End If
End Sub

Sub ProcessTransaction(t)
	Dim dTxn
	Dim bFound
	Dim sTmp, sMemo
	Dim vParts, sPart
	LogProgress Bcfg.IDString, "ProcessTransaction"
' from early 2008 the memo has a different length pattern
' where the :86: lines contain two 32-char segments separated by two spaces (as do their continuations)
' giving a total line length of 66
	Dim iTmp
	iTmp = InStr(t.Memo, Cfg.MemoDelimiter)
	If iTmp = 0 Then
		iTmp = Len(t.Memo)
	End If
	If iTmp > 32 Then
		sMemo = ""
		vParts = Split(t.Memo, Cfg.MemoDelimiter)
		For iTmp = LBound(vParts) To UBound(vParts)
			sPart = vParts(iTmp)
			If iTmp = LBound(vParts) And IsNumeric(Left(sPart, 1)) Then
				sPart = " " & sPart
			End If
			sTmp = Trim(Left(sPart, 32))
			If Len(sTmp) > 0 Then
				If Len(sMemo) > 0 Then
					sMemo = sMemo & Cfg.MemoDelimiter
				End If
				sMemo = sMemo & sTmp
			End If
			sTmp = Trim(Mid(sPart, 34))
			If Len(sTmp) > 0 Then
				sMemo = sMemo & Cfg.MemoDelimiter
				sMemo = sMemo & sTmp
			End If
		Next
		t.Memo = sMemo
	End If
	
	t.Payee = ABNA_FindPayee(t)
	If Cfg.ScriptDebugLevel > 5 Then
		MsgBox "Found payee: " & t.Payee
	End If
	dTxn = ABNA_FindTxnDate(t, bFound)
	If bFound Then
		If Cfg.ScriptDebugLevel > 5 Then
			MsgBox "Found Txn Date: " & dTxn
		End If
		t.TxnDateValid = True
		t.TxnDate = dTxn
	End If
	If NLG2EUR Then
		t.Amt = t.Amt / EURRate
	End If
End Sub

Sub EndSession()
	LogProgress Bcfg.IDString, "EndSession"
	If Cfg.TxnDumpFile <> "" Then
		DumpObjects Cfg.TxnDumpFile
	End If
End Sub

Function ABNA_FindPayee(t)
    Dim v
    Dim sPayee
    Dim iTmp, iTmp2
    Dim sMemo
    Dim aSepa, sInfo
    
    sMemo = t.FurtherInfo
    If sMemo = "" Then
        ABNA_FindPayee = sMemo
        Exit Function
    End If
    sPayee = GetSpecialPayee(sMemo)
    If sPayee <> "" Then
        ABNA_FindPayee = sPayee
        Exit Function
    End If
    v = Split(sMemo, Cfg.MemoDelimiter)
    If Not IsArray(v) Then
        ABNA_FindPayee = sMemo
        Exit Function
    End If
    sPayee = v(0)
    If Left(sPayee, 16) = "PROV.  TELEGIRO " Then
        sPayee = Mid(sPayee, 17)
    End If
    If Left(sPayee, 14) = "BETAALAUTOMAAT" or Left(sPayee, 4) = "BEA " Then
        sPayee = v(1)
        iTmp = InStr(sPayee, ",")
        If iTmp <> 0 Then
            sPayee = Trim(Left(sPayee, iTmp - 1))
        End If
    ElseIf IsNumeric(Left(sPayee, 1)) Then ' starts with account number
        sPayee = Trim(Mid(sPayee, 14))
        If sPayee = "" Then sPayee = v(1)
    ElseIf Left(sPayee, 5) = "GIRO " Then
        sPayee = Mid(sPayee, 6)
        While Len(sPayee) > 0 And (Left(sPayee, 1) = " " Or IsNumeric(Left(sPayee, 1)))
            sPayee = Mid(sPayee, 2)
        Wend
        sPayee = Trim(sPayee)
        If sPayee = "" Then
            sPayee = v(1)
        End If
    ElseIf Left(sPayee, 2) = "NI" And IsNumeric(Mid(sPayee, 3, 1)) Then
        sPayee = v(1)
    ElseIf Left(sPayee, 6) = "EC NR " Then
        sPayee = Trim(Mid(sPayee, 15))
        If sPayee = "" Then sPayee = v(1)
    ElseIf Left(sPayee, 3) = "EC " Then
        sPayee = Trim(Mid(sPayee, 12))
        If sPayee = "" Then sPayee = v(1)
    ElseIf Left(sPayee, 5) = "SEPA " Then
        If Left(v(3), 6) = "NAAM: " Then
            sPayee = Trim(Mid(v(3), 7))
        End If
    ElseIf Left(sPayee, 6) = "/TRTP/" Then
        sMemo = Replace(sMemo, Cfg.MemoDelimiter, "")
        aSepa = Split(Mid(sMemo,2), "/")
        For iTmp=2 To UBound(aSepa)-1
            Select Case aSepa(iTmp)
            Case "NAME"
                sPayee = aSepa(iTmp+1)
            Case "TRTP","PREF","NTRX","RTYP","MARF","EREF","IBAN","BIC","RTRN","SWOD"
            Case "REMI"
                sInfo = aSepa(iTmp+1)
            Case Else
                sInfo = sInfo & "/" & aSepa(iTmp+1)
            End Select
        Next
        t.FurtherInfo = sMemo
    End If
    ABNA_FindPayee = sPayee
End Function

':61:1302280228D300,N411NONREF
':86:SEPA PERIODIEKE OVERB.           IBAN: NL91AEGO0206533349
'BIC: AEGONL2U                    NAAM: C R SMALE
'OMSCHRIJVING: SPAREN

'SEPA PERIODIEKE OVERB.           IBAN: NL91AEGO0206533349;BIC: AEGONL2U                    NAAM: C R SMALE

':61:1303010301D88,N944NONREF
':86:SEPA IDEAL                       IBAN: NL44DEUT0496286323
'BIC: DEUTNL2N                    NAAM: NETGIRO PAYMENTS AB
'OMSCHRIJVING: DRWP1528296622 115 0000228999000 2 TICKETS TO BRUNO
'MARS TICKETMASTER NETHERLANDS   KENMERK: 01-03-2013 10:13 115000
'0228999000

Function ABNA_FindTxnDate(t, bFound)
    Dim sMemo
    Dim dTmp
    sMemo = t.FurtherInfo
    bFound = False
    If Left(sMemo, 15) = "BETAALAUTOMAAT " _
    Or Left(sMemo, 13) = "GELDAUTOMAAT " _
    Or Left(sMemo, 9) = "CHIPKNIP " Then
        dTmp = DateSerial(CInt(Mid(sMemo, 22, 2)) + 2000, _
            CInt(Mid(sMemo, 19, 2)), _
            CInt(Mid(sMemo, 16, 2)))
        dTmp = dTmp + TimeSerial(CInt(Mid(sMemo, 25, 2)), _
            CInt(Mid(sMemo, 28, 2)), _
            0)
    Elseif Left(sMemo, 4) = "BEA " _
        Or Left(sMemo, 4) = "GEA " _
        Or Left(sMemo, 5) = "CHIP " Then
            dTmp = DateSerial(CInt(Mid(sMemo, 25, 2)) + 2000, _
                CInt(Mid(sMemo, 22, 2)), _
                CInt(Mid(sMemo, 19, 2)))
            dTmp = dTmp + TimeSerial(CInt(Mid(sMemo, 28, 2)), _
                CInt(Mid(sMemo, 31, 2)), _
                0)
    Else
        Exit Function
    End If
    bFound = True
    ABNA_FindTxnDate = dTmp
End Function

Function GetSpecialPayee(sMemo)
    Dim sTmp
    Dim bDone
    Dim i
    Dim re
    Dim Matches
    
    Set re = New RegExp
    i = 1
    bDone = False
    sTmp = GetConfigString(IniSectionSpecialPayeeNames, _
        IniPatternPrefix & CStr(i), "")
    Do While sTmp <> "" And Not bDone
        re.Pattern = "^" & sTmp & ".*"
        Set Matches = re.Execute(sMemo)
        If Matches.Count > 0 Then
            sTmp = GetConfigString(IniSectionSpecialPayeeNames, _
                IniPayeePrefix & CStr(i), "")
            bDone = True
        Else
            i = i + 1
            sTmp = GetConfigString(IniSectionSpecialPayeeNames, _
                IniPatternPrefix & CStr(i), "")
        End If
    Loop
    GetSpecialPayee = sTmp
End Function

Function ABNA_FindServerTime(sFileIn)
    Dim iTmp
    Dim sTmp
    Dim iYear, iMon, iDay
    Dim dServer
' extract server date/time from file name!!
    iTmp = InStrRev(sFileIn, "\")
    If iTmp = 0 Then
        sTmp = sFileIn
    Else
        sTmp = Mid(sFileIn, iTmp + 1)
    End If
    If Left(sTmp, 5) = "MT940" And IsNumeric(Mid(sTmp, 6, 12)) Then
    	iYear = CInt(Mid(sTmp, 6, 2)) + 2000
    	iMon = CInt(Mid(sTmp, 8, 2))
    	iDay = CInt(Mid(sTmp, 10, 2))
' format changed around 1 August 2003 from DMY to YMD. This code is for the
' new situation!
        dServer = DateSerial(iYear, iMon, iDay)
        dServer = dServer + TimeSerial(CInt(Mid(sTmp, 12, 2)), CInt(Mid(sTmp, 14, 2)), CInt(Mid(sTmp, 16, 2)))
    Else
        dServer = Now
    End If
    ABNA_FindServerTime = dServer
End Function
