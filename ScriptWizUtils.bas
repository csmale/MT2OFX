Attribute VB_Name = "ScriptWizUtils"
Option Explicit

' CSV Field Codes
Public Const fldSkip = 0
Public Const fldAccountNum = 1
Public Const fldCurrency = 2
Public Const fldClosingBal = 3
Public Const fldAvailBal = 4
Public Const fldBookDate = 5
Public Const fldValueDate = 6
Public Const fldAmtCredit = 7
Public Const fldAmtDebit = 8
Public Const fldMemo = 9
Public Const fldBalanceDate = 10
Public Const fldAmount = 11
Public Const fldPayee = 12
Public Const fldTransactionDate = 13
Public Const fldTransactionTime = 14
Public Const fldChequeNum = 15
Public Const fldCheckNum = 15
Public Const fldFITID = 16
Public Const fldEmpty = 17  ' field is ignored but MUST be empty for recognition
Public Const fldBranch = 18
Public Const fldSign = 19   ' + or -
Public Const fldCategory = 20
Public Const fldPayeeCity = 21
Public Const fldPayeeState = 22
Public Const fldPayeeZip = 23
Public Const fldPayeeCountry = 24
Public Const fldPayeePhone = 25
Public Const fldPayeeAddress1 = 26
Public Const fldPayeeAddress2 = 27
Public Const fldPayeeAddress3 = 28
Public Const fldPayeeAddress4 = 29
Public Const fldPayeeAddress5 = 30
Public Const fldPayeeAddress6 = 31

