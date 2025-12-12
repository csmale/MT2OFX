VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSVOutputFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : CSVOutputFormatter
'    Project    : MT2OFX
'
'    Description: [type_description_here]
'
'    Modified   : $Author: Colin $ $Date: 13/06/11 9:04 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/CSVOutputFormatter.cls 1     13/06/11 9:04 Colin $"
' $History: CSVOutputFormatter.cls $
' 
' *****************  Version 1  *****************
' User: Colin        Date: 13/06/11   Time: 9:04
' Created in $/MT2OFX

'</CSCC>
Public QuoteChar As String
Public AlwaysQuoteFields As Boolean
Public UseBackslashEscape As Boolean

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       txtMakeLineDelimited
' Description:       Create CSV-compatible line from items
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       30/07/2004-23:32:57
'
' Parameters :       vFields (Variant)
'                    Delimiter (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function txtMakeLineDelimited(vFields As Variant, _
    Optional bBackslash As Boolean = False, _
    Optional bForceQuotes As Boolean = True, _
    Optional sDelim As String = ",", _
    Optional sQuote As String = """") As String
    
    Dim sTmp As String
    Dim sLine As String
    Dim i As Long
    If IsArray(vFields) Then
        For i = LBound(vFields) To UBound(vFields)
            sTmp = txtMakeFieldDelimited(CStr(vFields(i)), bBackslash, bForceQuotes, sDelim)
            If Len(sLine) > 0 Then
                sLine = sLine & sDelim
            End If
            sLine = sLine & sTmp
        Next
    Else
        sTmp = txtMakeFieldDelimited(CStr(vFields), bBackslash, bForceQuotes, sDelim)
        sLine = sTmp
    End If
    txtMakeLineDelimited = sLine
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       MakeFieldDelimited
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       COLIN-7
' Date-Time  :       13/06/2011-00:42:10
'
' Parameters :       sField (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function txtMakeFieldDelimited(sField As String, _
    Optional bBackslash As Boolean = False, _
    Optional bForceQuotes As Boolean = True, _
    Optional sDelim As String = ",", _
    Optional sQuote As String = """") As String
    Dim sRet As String
    sRet = sField
    If bBackslash Then
        sRet = Replace$(sRet, "\", "\\")          ' \ replaced with \\
        If Len(sQuote) = 1 Then
            sRet = Replace$(sRet, sQuote, "\" & sQuote)  ' " replaced with \"
        End If
    Else
        If Len(sQuote) = 1 Then
            sRet = Replace$(sField, sQuote, sQuote & sQuote)  ' " replaced with ""
        End If
    End If
    If bForceQuotes Or InStr(sRet, sDelim) > 0 Then
        sRet = sQuote & sRet & sQuote
    End If
    txtMakeFieldDelimited = sRet
End Function



