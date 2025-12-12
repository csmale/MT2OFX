VERSION 5.00
Begin VB.UserControl Hyperlink 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5115
   ScaleHeight     =   2385
   ScaleWidth      =   5115
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   1560
   End
   Begin VB.Label lblHyperlink 
      AutoSize        =   -1  'True
      Caption         =   "Hyperlink goes here (autosized)"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   2220
   End
End
Attribute VB_Name = "Hyperlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : Hyperlink
'    Project    : MT2OFX
'
'    Description: Hyperlink User Control
'
'    Modified   : $Author: Colin $ $Date: 2/11/05 23:03 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/Hyperlink.ctl 3     2/11/05 23:03 Colin $"
' $History: Hyperlink.ctl $
' 
' *****************  Version 3  *****************
' User: Colin        Date: 2/11/05    Time: 23:03
' Updated in $/MT2OFX
' V3.4 beta 1
'
' *****************  Version 2  *****************
' User: Colin        Date: 11/06/05   Time: 19:32
' Updated in $/MT2OFX
'</CSCC>

'Event Declarations:
Event Click() 'MappingInfo=lblHyperlink,lblHyperlink,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lblHyperlink,lblHyperlink,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lblHyperlink,lblHyperlink,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lblHyperlink,lblHyperlink,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Private Const clrLinkActive = vbBlue
Private Const clrLinkHot = vbRed
Private Const clrLinkInactive = vbBlack

Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long

Private Declare Function ScreenToClient Lib "User32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetDesktopWindow Lib "User32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub lblHyperlink_Click()
    RaiseEvent Click
   Dim sURL As String

  'open the URL using the default browser
   sURL = lblHyperlink.Caption

   Call RunShellExecute("open", sURL, 0&, 0&, SW_SHOWNORMAL)
End Sub

Private Sub Timer1_Timer()
   Dim pt As POINTAPI
   Dim X As Long
   Dim Y As Long
    
  'determine if the cursor is still over the link label
   With lblHyperlink
      GetCursorPos pt
      ScreenToClient UserControl.hWnd, pt
      X = pt.X * Screen.TwipsPerPixelX
      Y = pt.Y * Screen.TwipsPerPixelY
      If (X < 0) Or (X > .Width) Or _
         (Y < 0) Or (Y > .Height) Then
           'the cursor has moved outside, so
           'reset the label appearance
            lblHyperlink.ForeColor = clrLinkInactive
            lblHyperlink.FontUnderline = False
           'and disable the timer
            Timer1.Enabled = False
      End If
   End With
End Sub

Private Sub UserControl_Initialize()
    With lblHyperlink
        .ForeColor = clrLinkInactive
        .Left = 0
        .Width = UserControl.Width
        .Top = 0
        .Height = UserControl.Height
        .Caption = ""
    End With
End Sub

Private Sub RunShellExecute(sTopic As String, sFile As Variant, _
                           sParams As Variant, sDirectory As Variant, _
                           nShowCmd As Long)

  'execute the passed operation, passing
  'the desktop as the window to receive
  'any error messages
   Call ShellExecute(GetDesktopWindow(), _
                     sTopic, _
                     sFile, _
                     sParams, _
                     sDirectory, _
                     nShowCmd)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblHyperlink,lblHyperlink,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = lblHyperlink.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    lblHyperlink.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblHyperlink,lblHyperlink,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = lblHyperlink.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblHyperlink.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblHyperlink,lblHyperlink,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = lblHyperlink.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    lblHyperlink.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblHyperlink,lblHyperlink,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblHyperlink.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblHyperlink.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblHyperlink,lblHyperlink,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = lblHyperlink.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    lblHyperlink.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblHyperlink,lblHyperlink,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = lblHyperlink.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    lblHyperlink.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblHyperlink,lblHyperlink,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    lblHyperlink.Refresh
End Sub


Private Sub lblHyperlink_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
  'when the label is clicked, change
  'the label colour to indicate it's hot
   With lblHyperlink
      If .ForeColor = clrLinkActive Then
         .ForeColor = clrLinkHot
         .Refresh
      End If
   End With
End Sub

Private Sub lblHyperlink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
  'if not already highlighted, set the
  'label colour and start the timer to
  'poll for the mouse cursor position
   With lblHyperlink
      If .ForeColor = clrLinkInactive Then
         .ForeColor = clrLinkActive
         .FontUnderline = True
         Timer1.Interval = 100
         Timer1.Enabled = True
         .Refresh
      End If
   End With
End Sub

Private Sub lblHyperlink_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
  'mouse released, so restore the label
  'to clrLinkActive
   With lblHyperlink
      If .ForeColor = clrLinkHot Then
         .ForeColor = clrLinkActive
         .Refresh
      End If
   End With
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lblHyperlink.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    lblHyperlink.ForeColor = PropBag.ReadProperty("ForeColor", clrLinkInactive)
    lblHyperlink.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lblHyperlink.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblHyperlink.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    lblHyperlink.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    lblHyperlink.Caption = PropBag.ReadProperty("HyperLink", "Hyperlink goes here (autosized)")
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", lblHyperlink.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", lblHyperlink.ForeColor, clrLinkInactive)
    Call PropBag.WriteProperty("Enabled", lblHyperlink.Enabled, True)
    Call PropBag.WriteProperty("Font", lblHyperlink.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", lblHyperlink.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", lblHyperlink.BorderStyle, 0)
    Call PropBag.WriteProperty("HyperLink", lblHyperlink.Caption, "Hyperlink goes here (autosized)")
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblHyperlink,lblHyperlink,-1,Caption
Public Property Get HyperLink() As String
Attribute HyperLink.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    HyperLink = lblHyperlink.Caption
End Property

Public Property Let HyperLink(ByVal New_HyperLink As String)
    lblHyperlink.Caption() = New_HyperLink
    PropertyChanged "HyperLink"
End Property
