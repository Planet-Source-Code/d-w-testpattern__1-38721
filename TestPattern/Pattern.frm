VERSION 5.00
Begin VB.Form Pattern 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8220
   LinkTopic       =   "Pattern"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   268
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   548
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   7260
      Left            =   -270
      Picture         =   "Pattern.frx":0000
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   -525
      Visible         =   0   'False
      Width           =   9660
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   1260
         Top             =   450
      End
   End
End
Attribute VB_Name = "Pattern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Const SRCCOPY = &HCC0020

Private Sub MoveMouse(X As Single, Y As Single)
Dim Pt As POINTAPI
Pt.X = X
Pt.Y = Y
ClientToScreen hWnd, Pt
SetCursorPos Pt.X, Pt.Y
End Sub
Private Function Roll() As Integer

Dim Spot As POINTAPI
Dim Middle As Single
Middle = (Me.Height / Screen.TwipsPerPixelY) / 2
GetCursorPos Spot

If Spot.Y > Middle Then
Roll = (Spot.Y - Middle) \ 4
ElseIf Spot.Y < Middle Then
Roll = -(Middle - Spot.Y) \ 4
Else
Roll = 0
End If

If Abs(Roll) < 20 Then Roll = 0
'makes it sticky in the middle like
'vertical control on TV

End Function

Private Sub Form_Click()
End
End Sub

Private Sub Form_Load()
Top = 0
Left = 0
Height = Screen.Height
Width = Screen.Width
MoveMouse (Me.Width / Screen.TwipsPerPixelX) / 2, (Me.Height / Screen.TwipsPerPixelY) / 2
End Sub





Private Sub Timer1_Timer()

Dim Slow As Integer
Dim Offset As Integer
Static Sign As Integer
Static Motion As Integer

If Sgn(Roll) <> 0 Then
Sign = Sgn(Roll)
End If

Offset = (Screen.TwipsPerPixelY * 2)
Me.Cls
StretchBlt Me.hdc, 0, Motion, Me.ScaleWidth, Me.ScaleHeight, Picture1.hdc, 0, 0, 640, 480, SRCCOPY
'stretch 640x480 jpg to screen size, first picture

If Motion > 0 Then 'rolling up
StretchBlt Me.hdc, 0, Motion - (Me.ScaleHeight) - Offset, Me.ScaleWidth, Me.ScaleHeight, Picture1.hdc, 0, 0, 640, 480, SRCCOPY
Else 'rolling down
StretchBlt Me.hdc, 0, Motion + (Me.ScaleHeight) + Offset, Me.ScaleWidth, Me.ScaleHeight, Picture1.hdc, 0, 0, 640, 480, SRCCOPY
End If
'second picture above or below
'the bar in between pictures is just the black form
'showing between pictures

Slow = Roll
    
If Slow = 0 And Motion <> 0 Then Slow = 5 * Sign
'slow it down preparing to stop
Motion = Motion + Slow

If Abs(Motion) > Screen.Height \ Screen.TwipsPerPixelY Then
Motion = 0 'make it stop rolling
Slow = 0
End If

End Sub



