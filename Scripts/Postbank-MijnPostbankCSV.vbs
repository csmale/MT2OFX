' MT2OFX Input Processing Script for Postbank NL Asc/CSV formats
Option Explicit

Const ScriptVersion = "$Header: /MT2OFX/Postbank-MijnPostbankCSV.vbs 16    6/12/14 21:55 Colin $"

Const ScriptName = "Postbank-MijnPostbankCSV"
Const FormatName = "Postbank (NL) CSV (Mijn Postbank) Formaat - Post-SEPA"
Const ParseErrorMessage = "Kan regel niet ontleden."
Dim ParseErrorTitle : ParseErrorTitle = ScriptName
Const DebugRecognition = False

' fld#	len	inhoud
'	1	8	datum yyyymmdd
Const fldDatum = 1
'	2	32	omschrijving 1 = naam
Const fldNaam = 2
'	3	22	rekeningnummer (IBAN, display format with spaces)
Const fldRekNr = 3
'	4	10	tegenrekening (onbetrouwbaar)
Const fldTgnRekNr = 4
'	5	3	txn code XXX
Const fldTransCode = 5
'	6	2	Af/Bij
Const fldAfBij = 6
'	7	12	bedrag
Const fldBedrag = 7
'	8	12	mutatiesoort
Const fldMutatieSoort = 8
'	9	32	mededeling
Const fldMededeling = 9

Const BADatePat = " Pasvolgnr:\d{3} (\d{2})-(\d{2})-(\d{4}) (\d{2}):(\d{2}).*"
' Pasvolgnr:016 10-11-2014 16:26 Transactie:28R174 Term:JM0406

Const GMDatePat = " (\d{2})-(\d{2})-(\d{2}) (\d{2}):(\d{2}).*"
' dd-mm-yy 16:52 999X999  9999999 

Dim LastMemo	' last non-blank memo field seen

Sub Initialise()
    LogProgress ScriptVersion, "Initialise"
    If Not CheckVersion() Then
        Abort
    End If
End Sub

' function DescriptiveName
' returns a string with a descriptive name of this script
Function DescriptiveName()
	DescriptiveName = FormatName
End Function

Function ParseDate(sDate)	' expects dd-mm-yyyy or yyyymmdd
    Dim iYear, iMonth, iDay			' for dates
	If Mid(sDate, 3, 1) = "-" Then
' Old format = dd-mm-yyyy
		iDay = CInt(Left(sDate,2))
		iMonth = CInt(Mid(sDate,4,2))
		iYear = CInt(Mid(sDate,7,4))
	Else
' New format = yyyymmdd
		iYear = CInt(left(sDate,4))
    	iMonth = CInt(Mid(sDate,5,2))
    	iDay = CInt(Mid(sDate,7,2))
	End If
	ParseDate = DateSerial(iYear, iMonth, iDay)
End Function

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' returns False if someone else can have a go!
Function RecogniseTextFile()
    Dim vFields
    Dim sLine
    RecogniseTextFile = False
    sLine = ReadLine()
' must be CSV format
    vFields = ParseLineDelimited(sLine, ",", False)
    If TypeName(vFields) <> "Variant()" Then
        If DebugRecognition Then
            Message True, True, "not variant()", ScriptName
        End If
        Exit Function
    End If
    If UBound(vFields) <> 9 Then
        If DebugRecognition Then
            Message True, True, "wrong number of fields - " & UBound(vFields) & " instead of 9", ScriptName
        End If
        Exit function
    End If
    If vFields(fldDatum) <> "Datum" Or _
        vFields(fldNaam) <> "Naam / Omschrijving" Or _
        vFields(fldRekNr) <> "Rekening" Or _
        vFields(fldTgnRekNr) <> "Tegenrekening" Or _
        vFields(fldTransCode) <> "Code" Or _
        vFields(fldAfBij) <> "Af Bij" Or _
        Trim(vFields(fldBedrag)) <> "Bedrag (EUR)" Or _
        vFields(fldMutatieSoort) <> "MutatieSoort" Or _
        vFields(fldMededeling) <> "Mededelingen" Then
            If DebugRecognition Then
                Message True, True, "header mismatch", ScriptName
            End If
	   		Exit Function
    End If	   	

    LogProgress ScriptName, "File Recognised"
    RecogniseTextFile = True
End Function

Function LoadTextFile()
    Dim sLine       ' holds a line
    Dim sPat        ' holds the match pattern
    Dim vFields     ' array of fields in the line
    Dim sType       ' record type
    Dim sAcct       ' last account number
    Dim Stmt        ' holds the current statement
    Dim sTmp		  ' temporary string
    Dim sTmp2       ' another one
    Dim iTmp        ' temporary number
    Dim vDateBits	  ' parts of date
    Dim iYear		  ' year
    Dim sOms1, sOms2, sOms3, sOms4, sOms5	' description lines
    Dim iSeq		  ' transaction sequence number
		
    LoadTextFile = False
    sAcct = ""
    '	eat first (header) line
    If Not AtEof() Then
        sLine = ReadLine()
    End If
    Do While Not AtEOF()
        sLine = ReadLine()

' 20090307 CS: allow double quotes in the text by replacing the real delimiters with single quotes
        sLine = Replace(sLine, """,", "',")
        sLine = Replace(sLine, ",""", ",'")
        If Left(sLine, 1) = """" Then sLine = "'" & Mid(sLine, 2)
        If Right(sLine, 1) = """" Then sLine = Left(sLine, Len(sLine)-1) & "'"

        If Len(sLine) > 0 Then
' fix up double quotes within the fields which confuses the parser

            vFields = ParseLineDelimited(sLine, ",", False)
            If TypeName(vFields) <> "Variant()" Then
                MsgBox ParseErrorMessage, vbOKOnly+vbCritical, ParseErrorTitle
                Abort
                Exit Function
            End If
	' set up new transaction, and start a new statement if the account # changes
            If sAcct <> vFields(fldRekNr) Then
                Set Stmt = NewStatement()
                iSeq = 0
                sAcct = vFields(fldRekNr)
                Stmt.Acct = Replace(sAcct, " ", "")
                Stmt.BankName = "Postbank"
                Stmt.OpeningBalance.Ccy = "EUR"
                Stmt.OpeningBalance.BalDate = ParseDate(vFields(fldDatum))
                Stmt.ClosingBalance.Ccy = ""	' to force 00000000 as date in ledger bal
            End If
            NewTransaction
            iSeq = iSeq + 1
            LastMemo = ""
            Txn.Amt = ParseNumber(vFields(fldBedrag), ",")
            If vFields(fldAfBij) <> "Bij" Then	' could be "Af" or blank - also DR
                Txn.Amt = -Txn.Amt
            End If
	' this will put the NAME of the payee into Txn.Payee.
            Txn.Payee = vFields(fldNaam)
' CS 20101030 Make sure entire payee field gets into the Memo field
            Txn.Memo = Txn.Payee
            If vFields(fldTransCode) = "BA" Then
	' strip off trailing digits (often branch number)
	' comment out this line if required!
                Txn.Payee = TrimTrailingDigits(Txn.Payee)
            ElseIf vFields(fldTransCode) = "IC" Then
' CS 20141205 no longer needed
            ElseIf vFields(fldTransCode) = "ST" Then
                Txn.Payee = "Storting"
            End If
' CS 20101030 Truncate payee at ">"
            iTmp = InStr(Txn.Payee, ">")
            If iTmp > 0 Then Txn.Payee = Trim(Left(Txn.Payee, iTmp-1))
            Txn.ValueDate = ParseDate(vFields(fldDatum))
            Txn.BookDate = ParseDate(vFields(fldDatum))
            Stmt.ClosingBalance.BalDate = Txn.ValueDate
            Txn.IsReversal = False
            Txn.TxnType = TransType(vFields(fldTransCode), Txn.Amt)
            If vFields(fldTransCode) = "BA" Then
                vDateBits = ParseLineFixed(vFields(fldMededeling), BADatePat)
            ElseIf vFields(fldTransCode) = "GM" Then
                vDateBits = ParseLineFixed(vFields(fldMededeling), GMDatePat)
            Else
                vDateBits = False	' as long as it is not an array!
            End If
' the following code works for all cases as long as the date/time fields are in the right order
            If TypeName(vDateBits) = "Variant()" Then
                If UBound(vDateBits) >= 5 Then
                    Txn.TxnDate = DateSerial(vDateBits(3), CInt(vDateBits(2)), CInt(vDateBits(1))) _
                        + TimeSerial(CInt(vDateBits(4)), CInt(vDateBits(5)), 0)
                    Txn.TxnDateValid = True
                End If
            End If
' cash withdrawals - special string for payee
            If vFields(fldTransCode) = "GM" Or vFields(fldTransCode) = "PK" Then
                Txn.Payee = "Kasopname"
            End If

            Txn.Memo = Trim(vFields(fldMededeling))
        End If
    Loop
    LoadTextFile = True
End Function

Function TransType(PostbankCode, Amt)
	Select Case PostbankCode
	Case "AC"
		TransType = "DIRECTDEBIT"
	Case "BA"
		TransType = "POS"
	Case "CH"
		TransType = "CHECK"
	Case "GB"
		TransType = "CHECK"
	Case "GM"
		TransType = "ATM"
	Case "IC"
		TransType = "DIRECTDEBIT"
	Case "PK"
		TransType = "CASH"
	Case "ST"
		TransType = "DEP"
	Case "VZ"
		TransType = "DIRECTDEP"
	Case Else
		TransType = "OTHER"
	End Select
End Function
'	other known codes:
'	RES, TAN, GIN, GEW, DV, EUR, FL, GF, GT, OV, PO, TA
