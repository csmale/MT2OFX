' MT2OFX Input Processing Script for CAMT.053 ISO20022 XML Format

Option Explicit

Private Const ScriptVersion = "$Header$"

Const ScriptName = "CAMT053-XML"
Const FormatName = "ISO20022 CAMT.053 Format"
Const ParseErrorMessage = "Parse Error"
Dim ParseErrorTitle : ParseErrorTitle = ScriptName

Dim xDoc		' XML document being processed

Sub Initialise()
    LogProgress ScriptName, "Initialise"
    Set xDoc = Nothing
	If Not CheckVersion() Then
		Abort
	End If
End Sub

' function DescriptiveName
' returns a string with a descriptive name of this script
Function DescriptiveName()
	DescriptiveName = FormatName
End Function

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
	Dim bRet
	Dim xNode
	Dim xFields
	Dim i
	If xDoc Is Nothing Then
		Set xDoc = CreateObject("MSXML2.DOMDocument")
		xDoc.async = False
    End If
    bRet = xDoc.load(Session.FileIn)
    If bRet Then
        xDoc.setProperty "SelectionLanguage", "XPath"
        xDoc.setProperty "SelectionNamespaces", "xmlns:x='urn:iso:std:iso:20022:tech:xsd:camt.053.001.02'"
        bRet = (xDoc.selectNodes("/x:Document/x:BkToCstmrStmt").length > 0)
        If Not bRet Then
            MsgBox "Document is not a CAMT.053 file"
        End If
    End If
   ' namespace is "urn:iso:std:iso:20022:tech:xsd:camt.053.001.02"
   
	If bRet Then
		LogProgress ScriptName, "File Recognised"
	End If
	RecogniseTextFile = bRet
End Function

Function LoadTextFile()
    Dim xStmts ' statements in file
    Dim xStmt  ' current statement
    Dim xBals  ' list of balances
    Dim xBal   ' balance being processed
    Dim xTxns	' list of transactions to be processed
    Dim xTxn	' transaction being processed
    Dim sTmp, dAmt, sCcy, dtBal, sBal
    Dim i, j
    Dim Stmt
    
    Set xStmts = xDoc.selectNodes("/x:Document/x:BkToCstmrStmt/x:Stmt")
    For i=0 To xStmts.length-1
        Set xStmt = xStmts.item(i)
        Set Stmt = NewStatement()
        Stmt.Acct = xStmt.selectSingleNode("x:Acct/x:Id/x:IBAN").text
        Stmt.BankName = xStmt.selectSingleNode("x:Acct/x:Svcr/x:FinInstnId/x:BIC").text
        Stmt.StatementID = xStmt.selectSingleNode("x:Id").text
        Stmt.OpeningBalance.Ccy = xStmt.selectSingleNode("x:Acct/x:Ccy").text
        Stmt.ClosingBalance.Ccy = xStmt.selectSingleNode("x:Acct/x:Ccy").text
        Set xBals = xStmt.selectNodes("x:Bal")
        For j=0 To xBals.length-1
            Set xBal = xBals.item(j)
            dAmt = GetAmount(xBal, sCcy)
            dtBal = GetDateTimeChoice(xBal)
            sBal = xBal.selectSingleNode("x:Tp/x:CdOrPrtry/x:Cd").text
            Select Case sBal
            Case "PRCD"
                Stmt.OpeningBalance.Amt = dAmt
                Stmt.OpeningBalance.Ccy = sCcy
                Stmt.OpeningBalance.BalDate = dtBal
            Case "CLBD"
                Stmt.ClosingBalance.Amt = dAmt
                Stmt.ClosingBalance.Ccy = sCcy
                Stmt.ClosingBalance.BalDate = dtBal
            Case "CLAV"
                Stmt.AvailableBalance.Amt = dAmt
                Stmt.AvailableBalance.Ccy = sCcy
                Stmt.AvailableBalance.BalDate = dtBal
            End Select
        Next
        Set xTxns = xStmt.selectNodes("x:Ntry")
        For j=0 To xTxns.length-1
            Set xTxn = xTxns.item(j)
            If xTxn.selectSingleNode("x:Sts").text = "BOOK" Then
                NewTransaction
                Txn.Amt = GetAmount(xTxn, sCcy)
                Txn.Memo = xTxn.selectSingleNode("x:AddtlNtryInf").text
                Txn.BookDate = GetDateTimeChoice(xTxn.selectSingleNode("x:BookgDt"))
                Txn.ValueDate = GetDateTimeChoice(xTxn.selectSingleNode("x:ValDt"))
                Txn.BookingCode = xTxn.selectSingleNode("x:BkTxCd/x:Prtry/x:Cd").text
            End If
        Next
    Next
    LoadTextFile = True
End Function

Function GetDateTimeChoice(xNode)
    Dim xDT, sDT
    GetDateTimeChoice = NODATE
    If xNode Is Nothing Then
        Exit Function
    End If
    Set xDT = xNode.selectSingleNode("x:Dt")
    If xDT Is Nothing Then
        Set xDT = xNode.selectSingleNode("x:DtTm")
    End If
    If Not (xDT Is Nothing) Then
        sDT = Left(xDT.text, 10)
        GetDateTimeChoice = DateSerial(CInt(Left(sDT,4)), CInt(Mid(sDT,6,2)), CInt(Right(sDT,2)))
    End If
End Function

Function GetAmount(xNode, sCcy)
    Dim dAmt
    dAmt = ParseNumber(xNode.selectSingleNode("x:Amt").text, ".")
    If xNode.selectSingleNode("x:CdtDbtInd").text = "DBIT" Then
        dAmt = -dAmt
    End If
    sCcy = xNode.selectSingleNode("x:Amt").getAttribute("Ccy")
    GetAmount = dAmt
End Function
