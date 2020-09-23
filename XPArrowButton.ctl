VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl XPArrowButton 
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2565
   MaskColor       =   &H00000000&
   MaskPicture     =   "XPArrowButton.ctx":0000
   ScaleHeight     =   222
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   171
   ToolboxBitmap   =   "XPArrowButton.ctx":0702
   Begin PicClip.PictureClip pcButtonArrow 
      Left            =   240
      Top             =   1200
      _ExtentX        =   1905
      _ExtentY        =   1905
      _Version        =   393216
      Rows            =   3
      Cols            =   3
      Picture         =   "XPArrowButton.ctx":0A14
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1800
      Top             =   600
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   960
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "XPArrowButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************
'*                                                    *
'* Nom du fichier : XPArrowButton.ctl                 *
'* Programmeur : Francois Trudeau                     *
'* Creation : 11 Octobre 2004                         *
'* Revision : 30 Decembre 2004                        *
'*                                                    *
'******************************************************

Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINT_API) As Long

Private Type POINT_API
    X As Long
    Y As Long
End Type

Dim btButtonType As Byte
Dim Dot As POINT_API
Dim boolMousePressed As Boolean
Dim boolMouseIsOver As Boolean      'Track the current in and out mouse state
Dim boolMouseIsOverOld As Boolean   'Remember if mouse is over

Public Event Click()
Public Event MouseDown()
Public Event MouseOut()
Public Event MouseIn()


Public Enum ButtonTypeConst
    [Next (Green) ] = 0
    [Next (Blue) ] = 3
    [Back (Orange) ] = 6
End Enum

Private Sub UserControl_Initialize()
 
boolMousePressed = False
btButtonType = 0
boolMouseIsOver = False
boolMouseIsOverOld = False

End Sub

Private Sub UserControl_Paint()

DrawButton 0

End Sub

Private Sub UserControl_Resize()

DrawButton 0
UserControl.Height = pbButton.ScaleHeight
UserControl.Width = pbButton.ScaleWidth

End Sub

Public Sub DrawButton(btState As Byte)

pbButton.Picture = pcButtonArrow.GraphicCell(btState + btButtonType)
UserControl.PaintPicture pbButton.Picture, 0, 0

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

boolMouseIsOver = True

If Button = vbLeftButton Or Button = vbRightButton Then
    boolMousePressed = True
End If

If Not boolMousePressed Then
    DrawButton 1
End If

Timer1.Enabled = True

TrackMouse


End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

boolMousePressed = True
boolMouseIsOver = True

DrawButton 2
Timer1.Enabled = True

RaiseEvent MouseDown
UserControl.Parent.Refresh

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Timer1.Enabled = True
boolMousePressed = False

If MouseIsOver = True Then
    DrawButton 1
    
    RaiseEvent Click
    UserControl.Parent.Refresh
    
Else
    DrawButton 0
End If

End Sub

Private Sub UserControl_DblClick()

Call UserControl_MouseDown(vbLeftButton, 1, 1, 1)

End Sub

Private Sub CheckCurPos()

'UserControl.ScaleMode = 3 'must have this 'cause of x and y, to know how to calc
Call GetCursorPos(Dot) 'Get mouse position on the screen
ScreenToClient UserControl.hWnd, Dot  'Convert screen coordinates to client coordinates

End Sub

Private Function MouseIsOver() As Boolean

CheckCurPos

'Checking if mouse is over our control, by x and y
If Dot.X < UserControl.ScaleLeft Or _
    Dot.Y < UserControl.ScaleTop Or _
    Dot.X > (UserControl.ScaleLeft + UserControl.ScaleWidth - 1) Or _
    Dot.Y > (UserControl.ScaleTop + UserControl.ScaleHeight - 1) Then
    
    'Mouse is Out
    boolMouseIsOver = False
    MouseIsOver = False

Else
    'Mouse is In
    boolMouseIsOver = True
    MouseIsOver = True

End If

End Function

Private Sub Timer1_Timer()

TrackMouse

If Not MouseIsOver Then
    Timer1.Enabled = False
    DrawButton 0
    TrackMouse
End If

End Sub

Private Sub TrackMouse()

boolMouseIsOver = MouseIsOver

If boolMouseIsOver <> boolMouseIsOverOld Then
    
    If boolMouseIsOver = True Then
        RaiseEvent MouseIn
    Else
        RaiseEvent MouseOut
    End If
    
    UserControl.Parent.Refresh
    boolMouseIsOverOld = boolMouseIsOver
    
End If

End Sub

Public Property Get ButtonType() As ButtonTypeConst

ButtonType = btButtonType

End Property

Public Property Let ButtonType(ByVal vNewValue As ButtonTypeConst)

btButtonType = vNewValue

PropertyChanged ButtonType
DrawButton 0

End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

With PropBag
    btButtonType = .ReadProperty("ButtonType", 0)
End With

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   
With PropBag
    .WriteProperty "ButtonType", ButtonType, 0
End With

End Sub
