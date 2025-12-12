' MT2OFX Input Processing Script for French LA POSTE TSV formats

' This is the script for the French LA POSTE CCP accounts.
' It is written by (ahem...cloned and butchered from another
' example script) an American living In France so it Is
' documented in English.  Désolée pour vous les francophones.
'
' This script assumes that you have NOT altered the input file
' in any way.  Yes, one could delete transaction lines with no 
' great problem but it assumes a certain number of header lines
' and adding/deleteing these sorts of lines would have unpredictable
' results.
'
' This script was developed in the autumn of 2004 and represents
' LA POSTE's file format at that time.  No assumption should be
' made that this file format will never change.

Option Explicit
Private Const ScriptVersion = "$Header: /MT2OFX/La Poste-TSV.vbs 7     7/12/06 15:07 Colin $"
Const ScriptName = "La Poste"
Const FormatName = "Format La Poste (FR) TSV Téléchargé"
Const ParseErrorMessage = "Échec - fichier inconnu"
Dim ParseErrorTitle : ParseErrorTitle = ScriptName
Dim sTmp, sTmp2 ' holds temporary variables
Dim dDate, dDay, dMonth, dYear ' holds temporary variables for date processing
Dim Stmt        ' holds the statement object
Dim sQuickenCurrency : sQuickenCurrency = "" ' Use this value to fool Quicken
Dim vFields     ' hold an array of fields from the input line
Dim sLine       ' hold the input line
Dim i           ' holds generic integer counter
Dim MonthNames					' month names in dates
MonthNames = Array("janvier","fevrier","mars","avril","mai","juni","juillet","août","septembre","octobre","novembre","decembre")
' Either give the month names in an array as above or use SetLocale to get the
' system strings for the given locale. Otherwise the default locale will be used.
' The MonthNames array must have a multiple of 12 elements, which run from Jan-Dec in groups of
' 12, i.e. "Jan".."Dec","January".."December" etc. Lower/upper case is not significant.
SetLocale "fr-FR"

Dim sAcct     : sAcct = ""       ' hold the account number
Dim sAcctType : sAcctType = ""   ' hold the account type
Dim sAcctCcy  : sAcctCcy = ""    ' hold the account currency
Dim sAcctDate : sAcctDate = ""   ' hold the statement/balance date
Dim sAcctBal  : sAcctBal = ""    ' hold the account balance in the account currency
Dim sAcctBalF : sAcctBalF = ""   ' hold the account balance in Francs
Dim sQFI      : sQFI = "1668"    ' hold the Quicken Financial Institution Number we want to use to fool Quicken

' This subrountine is here for what ever purposes
' that MT2OFX needs.  Not currently used specifically
' for LA POSTE processing.
Sub Initialise()
    LogProgress ScriptName, "Initialise"
	If Not CheckVersion() Then
		Abort
	End If
' Initialise dictionary of month names
	InitialiseMonths MonthNames
End Sub

Function DescriptiveName()
	DescriptiveName = FormatName
End Function

' function RecogniseTextFile
' returns True if the input file is recognised by this script
' 
' For LA POSTE, the file is a TSV file of transactions with some
' header information...not a real statement.  This routine checks
' the header lines for expected values.  It also gathers some
' of the information for the Statement object here
Function RecogniseTextFile()
	Dim hdrLabel    ' if a header line, holds the Label field (1st)
	Dim hdrValue    ' if a header line, holds the Value field (2nd)

	'innocent until proven guilty...
 	RecogniseTextFile = False
	
	' This section reads through the header lines at the top of the
	' TSV file and harvests all the values it sees.  If it sees all
	' the values it expects to see then it populates the Statement
	' object.  It also positions to the beginning of the transaction
	' lines in the file.  Finding all the values that it expects to
	' find is the criteria for saying that this file is "recognized"
	' as a valid LA POSTE downloaded file.
	'
	' Note: this script was created with a single CCP account in euros
	'       if other types exist with different header configurations,
	'       this part of the script will need to be changed/scrapped.	
	For i = 1 To 7 'lines of input
		sLine = ReadLine()
		If AtEOF() then
			Exit For 'if the file is less than 6 lines long
		End If		
'Debugging...
'msgbox "#1" & vbcrlf & sline 
		vFields = ParseLineDelimited(sLine, vbTab)
		If ubound(vFields) <> 2 Then
			Exit For 'if there are less or more than 2 fields in any line
		End If	
		hdrLabel = vFields(1)
		hdrValue = vFields(2)
		' Parse which line we're looking at and process accordingly
		Select Case Trim(hdrLabel)
			Case "Num Compte", "Numéro Compte"
				sAcct = hdrValue
			Case "Libellé"
				sAcctType = hdrValue
			Case "Type"
				' CCP...?
			Case "Compte tenu en"
				' If importing into Quicken and this value NEEDS to be
				' something specific, set it here
				If sQuickenCurrency <> "" Then
					sAcctCcy = sQuickenCurrency
				Else
					sAcctCcy = hdrValue
					' Parse the currency name and change it into the value for the output file
					' Add currencies to this SELECT statement if necessary
					Select Case sAcctCcy
						Case "euros"
							sAcctCcy = "EUR"
						Case Else
							sTmp = "Il faut ajouter cette nouvelle sorte de monnaie: " & _
							       sAcctCcy & " dans le script pour La Poste"
							msgbox(sTmp)	
							sAcctCcy = "" ' Re-initialize to indicate that we've not recognized the file
					End Select
				End If	
			Case "Date"
	MsgBox "Date: " & hdrValue
				If InStr(hdrValue, "/") = 0 Then
					sAcctDate = ParseDateEx(hdrValue, "DMY", " ")
				Else
					sAcctDate = ParseDateEx(hdrValue, "DMY", "/")
				End If
			Case "Solde (EUROS)"
				sAcctBal = hdrValue
			Case "Solde (FRANCS)"
				sAcctBalF = hdrValue
			Case Else
				Exit For
		End Select
	Next

	' Skip and look for header line
	If Not AtEOF() Then
		' Read the blank line
		If Len(Trim(sLine)) > 0 then
			sLine = ReadLine()
		End if
		If Not AtEOF() Then
			' Read the header line
'Debugging...
' msgbox "#2" & vbCrlf & sline 
			sLine = ReadLine()
			If Not AtEOF() Then
				' Check to see if this is really the header line
				' Note: If the account is not in Euros, then this
				'       will need to be modified.
'Debugging...
' msgbox "#3" & vbcrlf & sLine 
				vFields = ParseLineDelimited(sLine, vbTab)
				If UBound(vFields) = 4 Then
					If vFields(1) = "Date" And _
					   vFields(2) = "Libellé" And _
					   ((vFields(3) = "Montants(EUROS)") Or (vFields(3) = "Montant(EUROS)")) And _
					   ((vFields(4) = "Montants (FRANCS)") Or (vFields(4) = "Montant(FRANCS)")) Then
					      ' Guilty as charged!
						RecogniseTextFile = True
					End If
				End If
			End If			
		End If
	End If
End Function

'function LoadTextFile
'This function is the main processing loop
'
'It loops through the input file and reads through the
'transaction records.  It determines the transaction
'type from the verbiage in the transaction description
'and whether it's a credit or debit.
'
'Note: the input file is "re-wound" to the beginning 
'      after called RecogniseTextFile so we have to
'      do some priming reads at the beginning to get
'      to the transaction records.
Function LoadTextFile()
Dim sBookDate ' holds the booking date from each line
Dim sTxnDate  ' holds the transaction date if the line contains one
Dim sTxnAmt   ' holds the transaction amount
Dim sPayee    ' holds the transaction payee field
Dim sTxnType  ' holds the transaction type value
Dim sMemo     ' holds any memo value
Dim sCheckNum ' holds a check number if there is one
Dim sDesc	  ' original txn description

      'See if there is a Quicken Financial Institution Code to insert.
      'If so, put it in the BCFG object to be output in the OFX/QFX file.
      If sQFI <> "" Then
		BCfg.IntuitBankID = sQFI
      End If 

'msgbox "IDString" & bcfg.IDString & vbCrlf & _
'       "Structured86" & bcfg.Structured86 & vbCrlf & _
'       "SkipEmptyMemoFields" & bcfg.SkipEmptyMemoFields & vbCrlf & _
'       "FindPayeeFn" & bcfg.FindPayeeFn & vbCrlf & _
'       "FindTxnDateFn" & bcfg.FindTxnDateFn & vbCrlf & _
'       "FindServerTimeFn" & bcfg.FindServerTimeFn & vbCrlf & _
'       "IntuitBankID" & bcfg.IntuitBankID & vbCrlf & _
'       "ScriptFile" & bcfg.ScriptFile

      'Again... innocent until proven guilty
      LoadTextFile = False
	
	' Create the Statement Object and populate a few fields
	Set Stmt = NewStatement()
	Stmt.Acct = sAcct
	Stmt.BankName = "La Poste"
	Stmt.StatementID = MakeGUID()
	Stmt.ClosingBalance.CCY = sAcctCcy
	Stmt.ClosingBalance.Amt = ParseNumber(sAcctBal, ",")
	Stmt.ClosingBalance.BalDate = sAcctDate
	
	'Priming reads to get to transaction lines
	For i = 1 To 9
		sLine = ReadLine()
	Next
	
	'Loop through the TSV file - main processing
	Do While Not AtEOF()
'Debugging...
'msgbox "#4" & vbcrlf & sline 
		'If this line really has anything in it...
		If Len(sLine) > 0 Then
			'...then it should have at least one <tab> in it
			vFields = ParseLineDelimited(sLine, vbTab)
			'...otherwise a debugging message anad get out
			If TypeName(vFields) <> "Variant()" Then
				MsgBox ParseErrorMessage & vbCrlf & "In: LoadTextFile()" & vbCrlf & "sLine:" & vbCrlf & sLine, vbOkonly+vbCritical, ParseErrorTitle
				Abort
				Exit Function
			End If

'Debugging...
'msgbox sline 
'& vbCrlf & vfields(1) & vbCrlf & vfields(2) & vbCrlf & vfields(3) & vbCrlf & vfields(4)

			'Let's initialize those values for this record/line
			sBookDate = ""
			sTxnDate = ""
			sTxnAmt = ""
			sPayee = ""
			sDesc = vFields(2) 'Store the original transaction description
			sMemo = sDesc
			sCheckNum = ""
			sTxnType = "CREDIT"
			If instr(vFields(3),"-") = 1 Then
				sTxnType = "DEBIT"
			End If
			
			'Store the booking date
			' (Oh Boy, the "Year 2100 Problem!!!")
		    sBookDate = dateserial("20" & right(vFields(1),2),mid(vFields(1),4,2),left(vFields(1),2))
			
			'Reformat the amount field for left side of the Atlantic tastes
			' 20061109 CS leave that for ParseNumber
			sTxnAmt = Trim(vFields(3))
			
			'Here are the "If/ElseIf" statements to check the Payee field
			'and determine a transaction type from that.  It also allows
			'one to strip some unwanted info/translate it/etc. for 
			'specific transaction types.
			
			'Visa Card Purchases
			If StartsWith(sDesc, "ACHAT CB ") Then
				sPayee = replace(sDesc,"ACHAT CB ","")
				sTmp = right(sPayee,8)
				sTxnDate = dateserial("20" & right(sTmp,2),mid(sTmp,4,2),left(sTmp,2)) 
				sPayee = left(sPayee,len(sPayee)-9)
				sMemo = "Visa Card Purchase"
				
			'Check/Cheque Entries	
			Elseif StartsWith(sDesc,"CHEQUE N ") Then
				sPayee = replace(sDesc,"CHEQUE N ","")
				sCheckNum = right(sPayee,7)
				sPayee = "**Check Payee**"
				sTxnType = "CHECK"
				sMemo = "Check Purchase"
			Elseif StartsWith(sDesc,"CHEQUE N. ") Then
				sPayee = replace(sDesc,"CHEQUE N. ","")
				sCheckNum = right(sPayee,7)
				sPayee = "**Check Payee**"
				sTxnType = "CHECK"
				sMemo = "Check Purchase"

			'Standing Orders				
			Elseif StartsWith(sDesc, "VIREMENT PERMANENT") Then
				sPayee = "**Standing Order**"
				sTxnType = "REPEATPMT"
				sMemo = "Standing Order"

			'"Prélevements"/Direct Debits	
			Elseif StartsWith(sDesc, "PREL DE ") Then
				sPayee = replace(sDesc,"PREL DE ","")
				sTxnType = "DIRECTDEBIT"
				
			'ATM Withdrawals	
			Elseif StartsWith(sDesc,"CARTE VISA     ") Then
				sTxnDate = dateserial("20" & mid(sDesc,22,2),mid(sDesc,19,2),mid(sDesc,16,2)) 
				sPayee = replace(sDesc,"CARTE VISA     ","ATM on: ")
				sTxnType = "ATM"
				
			'Point Of Sale (POS) transactions...I think!	
			Elseif StartsWith(sDesc,"RETRAIT ") Then
				sTmp = right(sDesc,8)
				sTxnDate = dateserial("20" & right(sTmp,2),mid(sTmp,4,2),left(sTmp,2)) 
				sPayee = rtrim(left(sDesc,len(sDesc)-9))
				sTxnType = "POS"
				
			'Transfer of funds INTO the account	
			Elseif StartsWith(sDesc,"VIR. DE ") Then
				sPayee = replace(sDesc,"VIR. DE ","Transfer From: ")
                sTxnType = "DIRECTDEP"
			Elseif StartsWith(sDesc,"VIREMENT DE ") Then
				sPayee = replace(sDesc,"VIREMENT DE ","Transfer From: ")
                sTxnType = "DIRECTDEP"
				
			'Account and Card fees	
			Elseif StartsWith(sDesc,"COTISATION ") Or StartsWith(sDesc, "FRAIS DE VIREMENT") Then
				sTxnType = "FEE"

			'Payment via TIP	
			Elseif StartsWith(sDesc,"TIP DE ") Then
				sPayee = replace(sDesc,"TIP DE ","")
				sTxnType = "DIRECTDEBIT"
					
                  'Payment to another account	
			      Elseif StartsWith(sDesc,"VIREMENT POUR   ") Then
					  sPayee = replace(sDesc,"VIREMENT POUR   ","Payment To: ")
					  sTxnType = "PAYMENT"
					
                  'Cash deposit	
			      Elseif InStr(sDesc,"VERSEMENT EFFECTUE") > 0 Then
					  sPayee = "Cash Deposit"
					  sTxnType = "DEP"
				  
				  'If this is another transaction type not yet covered
				  Else
				  sPayee = "**Unknown Txn Type** " & sDesc
				sMemo = sDesc
			End If

			'Ok we're ready to create a new transaction object
			'and give it the values we've just gathered
			NewTransaction()

'More debugging opportunities			
'msgbox sline & vbCrlf & _ 
'	"check " & schecknum & vbCrlf & _
'	"amount " & stxnamt & vbCrlf & _
'	"txn date " & stxndate & vbCrlf & _
'	"book date " & sbookdate & vbCrlf & _
'	"memo " & smemo & vbCrlf & _
'	"payee " & spayee & vbCrlf & _
'	"txn type " & stxntype

			Txn.FITID = MakeGUID()
			Txn.CheckNum = sCheckNum
			Txn.Amt = ParseNumber(sTxnAmt, ",")
			If sTxnDate <> "" Then
				Txn.TxnDate = sTxnDate
				Txn.TxnDateValid = True
			End If			
			Txn.BookDate = sBookDate
			Txn.ValueDate = sBookDate
			Txn.FurtherInfo = sMemo
			Txn.Payee = sPayee
			Txn.TxnType = sTxnType
		End If
		sLine = ReadLine()
	Loop
	
	'We be done...
	LoadTextFile = True
End Function