Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/SparkasseHildesheim.vbs 1     19/04/08 22:20 Colin $"
Private Const FormatName = "Sparkasse Hildesheim MT940-format."
Const ScriptName = "SparkasseHildesheim"

Dim QuickenBankID

' Property List is an array of arrays, each of which has the following elements:
'	1. Property key - used to reference properties
'	2. Property name - used as a label in the config screen
'	3. Property description - used as a description or tooltip in the config screen
'	4. Data type - ptString, ptBoolean, ptInteger, ptFloat, ptDate, ptChoice
'	5. Value list (will be displayed in a combobox) - array of values (Only with ptChoice)
Dim aPropertyList
aPropertyList = Array( _
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
	QuickenBankID = GetProperty("QuickenBankID")
	Bcfg.IntuitBankID = QuickenBankID
	Session.ServerTime = FindServerTime(Session.FileIn)
End Sub

Sub ProcessStatement(s)
	LogProgress Bcfg.IDString, "ProcessStatement"
End Sub

Sub ProcessTransaction(t)
	Dim dTxn
	Dim bFound
	Dim sTmp, sMemo
	LogProgress Bcfg.IDString, "ProcessTransaction"
	
	t.Payee = FindPayee(t)
	If Cfg.ScriptDebugLevel > 5 Then
		MsgBox "Found payee: " & t.Payee
	End If
	dTxn = FindTxnDate(t, bFound)
	If bFound Then
		If Cfg.ScriptDebugLevel > 5 Then
			MsgBox "Found Txn Date: " & dTxn
		End If
		t.TxnDateValid = True
		t.TxnDate = dTxn
	End If
End Sub

Sub EndSession()
	LogProgress Bcfg.IDString, "EndSession"
	If Cfg.TxnDumpFile <> "" Then
		DumpObjects Cfg.TxnDumpFile
	End If
End Sub

Function FindPayee(t)
    Dim v
    Dim sPayee
    Dim iTmp
    Dim sMemo
    
    sMemo = t.FurtherInfo
    If sMemo = "" Then
        FindPayee = sMemo
        Exit Function
    End If
    v = Split(sMemo, Cfg.MemoDelimiter)
    If Not IsArray(v) Then
        FindPayee = sMemo
        Exit Function
    End If
    sPayee = v(0)
    FindPayee = sPayee
End Function

Function FindTxnDate(t, bFound)
    Dim sMemo
    Dim dTmp
    sMemo = t.Memo
    
    bFound = False
    FindTxnDate = dTmp
End Function

Function FindServerTime(sFileIn)
    Dim iTmp
    Dim sTmp
    Dim iYear, iMon, iDay
    Dim dServer

    FindServerTime = dServer
End Function
