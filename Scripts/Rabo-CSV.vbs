' MT2OFX Input Processing Script for Rabobank NL CSV format

Option Explicit

Private Const ScriptVersion = "$Header: /MT2OFX/Rabo-CSV.vbs 1     26/11/08 0:04 Colin $"

Const ScriptName = "Rabo-CSV"
Const FormatName = "Rabobank Nederland CSV Formaat"
Dim NoOFXMessage : NoOFXMessage = "Van dit bestandstype kunt u geen OFC of OFX produceren omdat het geen saldoinformatie bevat." _
	& vbCrLf & vbCrLf & "Kies een ander uitvoerformaat zoals QIF."

' Veld Lengte Veldnaam   Inhoud/Opmerking
' 1    X(10)  VAN_REK    Eigen rekeningnummer
' 2    X(3)   MUNTSOORT  EUR
' 3    9(8)   RENTEDATUM Formaat: EEJJMMDD
' 4    X(1)   BY_AF_CODE C of D
' 5    9(14)  BEDRAG     2 decimalen, na scheidingsteken "." op positie 12
' 6    X(10)  NAAR_REK   Tegenrekeningnummer
' 7    X(24)  NAAR_NAAM
' 8    9(8)   BOEKDATUM  Formaat: EEJJMMDD
' 9    X(2)   BOEKCODE
' 10   X(6)   BUDGETCODE
' 11   X(32)  OMSCHR1
' 12   X(32)  OMSCHR2
' 13   X(32)  OMSCHR3
' 14   X(32)  OMSCHR4
' 15   X(32)  OMSCHR5
' 16   X(32)  OMSCHR6
