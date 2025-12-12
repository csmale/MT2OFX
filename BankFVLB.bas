Attribute VB_Name = "BankFVLB"
Option Explicit

' $Header: /MT2OFX/BankFVLB.bas 2     4/04/04 21:24 Colin $

Public Function FVLB_FindPayee(t As Txn) As String
    Dim v As Variant
    Dim sPayee As String
    Dim iTmp As Integer
    Dim sMemo As String

    sPayee = t.Str86.Payee
    If sPayee = "" Then sPayee = t.Str86.GetField(sfBookingText)
    If sPayee = "" Then sPayee = "?"
    FVLB_FindPayee = sPayee
End Function
