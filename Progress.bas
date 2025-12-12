Attribute VB_Name = "Progress"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : Progress
'    Project    : MT2OFX
'
'    Description: Progress Bar Control
'
'    Modified   : $Author: Colin $ $Date: 6/10/09 0:37 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/Progress.bas 9     6/10/09 0:37 Colin $"
' $History: Progress.bas $
' 
' *****************  Version 9  *****************
' User: Colin        Date: 6/10/09    Time: 0:37
' Updated in $/MT2OFX
' quiet mode support for watcher
'
' *****************  Version 7  *****************
' User: Colin        Date: 20/04/08   Time: 10:06
' Updated in $/MT2OFX
' For 3.5 beta 1
'
' *****************  Version 6  *****************
' User: Colin        Date: 7/12/06    Time: 15:07
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 3  *****************
' User: Colin        Date: 2/11/05    Time: 23:03
' Updated in $/MT2OFX
' V3.4 beta 1
'
' *****************  Version 2  *****************
' User: Colin        Date: 11/06/05   Time: 19:33
' Updated in $/MT2OFX
'</CSCC>

Private frmMyProgress As frmProgress

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       InitProgress
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       31/05/2005-22:37:26
'
' Parameters :       sPhase (String)
'                    iMax (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub InitProgress(sPhase As String, ByVal iMin As Long, ByVal iMax As Long)
' 20090917 CS: quiet mode - no gui feedback
    If CmdParams.Quiet Then
        Exit Sub
    End If
    If frmMyProgress Is Nothing Then
        Set frmMyProgress = New frmProgress
    End If
    With frmMyProgress
        .lblPhase.Caption = sPhase
        If iMin < 0 Then
            .lMin = 0
        Else
            .lMin = iMin
        End If
        If iMax > .lMin Then
            .lMax = iMax
        Else
            .lMax = .lMin + 1
        End If
        If .lMax > 500000 Then
            .Factor = 100
        Else
            .Factor = 1000
        End If
        .ResetProgress
        .SetProgress iMin
    End With
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ShowProgress
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       31/05/2005-22:40:55
'
' Parameters :       iVal (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub ShowProgress(iVal As Long)
    DoEvents
    If frmMyProgress Is Nothing Then
'        Debug.Assert False
        Exit Sub
    End If
    With frmMyProgress
        .SetProgress iVal
    End With
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CloseProgress
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       31/05/2005-22:42:46
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub CloseProgress()
    If Not frmMyProgress Is Nothing Then
        Unload frmMyProgress
        Set frmMyProgress = Nothing
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       SetProgressTopmost
' Description:       Manipulate Topmost setting on progress window to give other windows a chance
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       8/15/2006-17:02:42
'
' Parameters :       bTopmost (Boolean)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function SetProgressTopmost(bTopmost As Boolean) As Boolean
    If frmMyProgress Is Nothing Then
        SetProgressTopmost = False
    Else
        SetProgressTopmost = frmMyProgress.SetTopmost(bTopmost)
    End If
End Function
