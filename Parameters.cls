VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FromClipboard As Boolean
Public InputFile As String
Public OutputFile As String
Public ScriptName As String
Public OutputType As String
Public Quiet As Boolean

Private Sub Class_Initialize()
    FromClipboard = False
    InputFile = ""
    OutputFile = ""
    ScriptName = ""
    OutputType = ""
    Quiet = False
End Sub
