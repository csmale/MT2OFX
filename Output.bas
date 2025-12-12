Attribute VB_Name = "Output"
Option Explicit

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Output
' Description:       [type_description_here]
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       14/05/2010-13:37:04
'
' Parameters :       s (Session)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function Output(s As Session) As Boolean
    If PrepareSession(s) Then
        If OutputGeneric(s) Then
            Output = OutputSpecific(s)
        Else
            Output = False
        End If
    Else
        Output = False
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       PrepareSession
' Description:       [type_description_here]
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       14/05/2010-13:28:16
'
' Parameters :       s (Session)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function PrepareSession(s As Session) As Boolean
    PrepareSession = False
    If s.FileIn = "" Then
        Exit Function
    End If
    s.PayeeMapFile = Cfg.PayeeMapFile
    s.PayeeMapIgnoreCase = Cfg.PayeeMapIgnoreCase
    s.ServerTime = NODATE
    s.Language = "ENG"
    PrepareSession = True
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OutputGeneric
' Description:       [type_description_here]
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       14/05/2010-13:30:57
'
' Parameters :       s (Session)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function OutputGeneric(s As Session) As Boolean
    Dim st As Statement
    Dim t As Txn
    For Each st In s.Statements
' get statement ID

        For Each t In st.Txns
        
        Next
    Next
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OutputSpecific
' Description:       [type_description_here]
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       14/05/2010-13:31:22
'
' Parameters :       s (Session)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function OutputSpecific(s As Session) As Boolean
    Select Case s.OutputFileType
    Case FileFormatOFC
        OutputSpecific = OutputOFC(s)
    Case FileFormatOFX, FileFormatQFX
        OutputSpecific = OutputOFX(s)
    Case FileFormatOFC
        OutputSpecific = OutputOFC(s)
    Case FileFormatQIF
        OutputSpecific = OutputQIF(s)
    Case FileFormatCustom
        OutputSpecific = OutputGeneric(s)
    Case Else
        OutputSpecific = False
    End Select
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OutputOFX
' Description:       [type_description_here]
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       14/05/2010-13:35:23
'
' Parameters :       s (Session)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function OutputOFX(s As Session) As Boolean

End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OutputOFC
' Description:       [type_description_here]
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       14/05/2010-13:35:07
'
' Parameters :       s (Session)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function OutputOFC(s As Session) As Boolean

End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OutputQIF
' Description:       [type_description_here]
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       14/05/2010-13:36:02
'
' Parameters :       s (Session)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function OutputQIF(s As Session) As Boolean

End Function

#If False Then
Init prepare
server Date
for each stmt
    statement dates etc
    stmt.id
    for each txn
        dates
        FITID
        Payee mapping
        Case correction
    Next Txn
    balances, totals etc
Next stmt

OUTPUT (format specific)
Init Output
File header
for each stmt
    Statement header
    for each txn
        transaction
    Next Txn
    Statement trailer
Next stmt
File trailer

custom Output
session object populated
do all the PREPARE actions
then OUTPUT

OUTPUT_Custom
CallScriptFunction Initialise
callScriptFunction DoCustomOutput(Session s)
CallScriptFunction Terminate

#End If

