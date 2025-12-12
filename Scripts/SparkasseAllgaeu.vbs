Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/SparkasseAllgaeu.vbs 2     2/01/08 0:53 Colin $"
Private Const FormatName = "Sparkasse Allgaeu MT940-formaat."
Const ScriptName = "SparkasseAllgaeu"

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
End Sub

Sub ProcessStatement(s)
	LogProgress Bcfg.IDString, "ProcessStatement"
End Sub

'Public Const sfBookingText = 0
'Public Const sfBatchNum = 10
'Public Const sfDetails0 = 20
'Public Const sfDetailsLast = 29
'Public Const sfBankPayee = 30
'Public Const sfAcctPayee = 31
'Public Const sfNamePayee = 32
'Public Const sfNamePayee2 = 33
'Public Const sfTextCodeSupplement = 34
'Public Const sfExtra0 = 60
'Public Const sfExtraLast = 63

Sub ConcatMemo(t, s)
	If s = "" Then
		Exit Sub
	End If
	If Len(t.Memo) > 0 Then
		t.Memo = t.Memo & Cfg.MemoDelimiter
	End If
	t.Memo = t.Memo & s
End Sub


Sub ProcessTransaction(t)
	Dim dTxn
	Dim bFound
	Dim i
	Dim sTmp
	LogProgress Bcfg.IDString, "ProcessTransaction"

	t.Memo = ""
	For i=sfDetails0 To sfDetailsLast
		sTmp = Trim(t.Str86.GetField(CLng(i)))
		ConcatMemo t, sTmp
	Next

	If t.Amt < 0 Then
		t.TxnType = "DEBIT"
	Else
		t.TxnType = "CREDIT"
	End If

	t.Payee = FindPayee(t)
	If Cfg.ScriptDebugLevel > 5 Then
		MsgBox "Found payee: " & t.Payee
	End If
	
	Select Case t.Str86.GetField(sfBookingText)
	Case "BAR"
		t.Payee = "Cash Withdrawal"
		t.TxnType = "ATM"
	End Select
	
' statement numbers are unreliable se generate a FITID from the statement Date
	dTxn = t.Statement.OpeningBalance.BalDate
	t.FITID = CStr(Year(dTxn)) _
		& "." & Right("000" & CStr(DatePart("y", dTxn)), 3) _
		& "." & CStr(t.Index)
End Sub

Sub EndSession()
	LogProgress Bcfg.IDString, "EndSession"
'	DumpObjects "C:\dump.txt"
	If Cfg.TxnDumpFile <> "" Then
		DumpObjects Cfg.TxnDumpFile
	End If
End Sub

Function FindPayee(t)
    Dim sPayee
    sPayee = t.Str86.GetField(sfNamePayee)
    FindPayee = sPayee
End Function

