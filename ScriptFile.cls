VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScriptFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : ScriptFile
'    Project    : MT2OFX
'
'    Description: Handles Script Files
'
'    Modified   : $Author: Colin $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/ScriptFile.cls 5     15/11/10 0:20 Colin $"
' $History: ScriptFile.cls $
' 
' *****************  Version 5  *****************
' User: Colin        Date: 15/11/10   Time: 0:20
' Updated in $/MT2OFX
'
' *****************  Version 4  *****************
' User: Colin        Date: 15/06/09   Time: 19:25
' Updated in $/MT2OFX
' For transfer to new laptop
'
' *****************  Version 2  *****************
' User: Colin        Date: 7/12/06    Time: 15:07
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 1  *****************
' User: Colin        Date: 8/09/06    Time: 23:56
' Created in $/MT2OFX

'</CSCC>
Private Const csScriptSubDir As String = "Scripts"

Private fso As New FileSystemObject

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Find
' Description:       Find a script file in various directories
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       9/8/2006-23:53:36
'
' Parameters :       sScript (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function Find(sScript As String) As String
    Dim sPath As String
    If PathIsRelative(sScript) Then
        sPath = Cfg.ScriptPath & "\" & sScript
        If Not fso.FileExists(sPath) Then
            sPath = Cfg.AppPath & "\" & sScript
            If Not fso.FileExists(sPath) Then
                sPath = Cfg.CustomOutputPath & "\" & sScript
                If Not fso.FileExists(sPath) Then sPath = ""
            End If
        End If
    Else
        sPath = sScript
        If Not fso.FileExists(sPath) Then sPath = ""
    End If
    Debug.Assert Len(sPath) > 0
    Find = sPath
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetContents
' Description:       Return the contents of a script file as a big long string
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       9/8/2006-23:54:02
'
' Parameters :       sScript (String)
'--------------------------------------------------------------------------------
'</CSCM>
' 20090310 CS: now uses InputFile to read script for Unicode support
Public Function GetContents(sScript As String) As String
    Dim sTmp As String
    Dim sPath As String
    Dim xFile As New InputFile
    On Error GoTo bad_open
    sPath = Find(sScript)
    If sPath = "" Then
' file not found: OpenTextFile will raise the error and log it
        sPath = sScript
    End If
    If Not xFile.OpenFile(sPath) Then
        GoTo bad_open
    End If
    sTmp = xFile.EntireFile
    xFile.CloseFile
    GetContents = sTmp
go_back:
    Exit Function
bad_open:
    LogMessage True, True, _
        "Unable to open script file " & sPath & ": " & Err.Description, _
        "MT2OFX Script Error"
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume go_back
End Function



