VERSION 5.00
Begin VB.Form Form16 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form16"
   ClientHeight    =   10035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15120
   LinkTopic       =   "Form16"
   Picture         =   "Form16.frx":0000
   ScaleHeight     =   10035
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Image Image1 
      Height          =   10050
      Left            =   0
      Picture         =   "Form16.frx":0C97
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15120
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Sub Form_Load()
    Dim Rec As RECT, hRgn As Long
    GetWindowRect Me.hWnd, Rec
    hRgn = CreateRoundRectRgn(0, 0, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, 50, 50) '这里的两个50是分别用来设置圆角的宽和高的
    SetWindowRgn Me.hWnd, hRgn, True
End Sub
Private Sub Image1_Click()
Form5.Visible = True
End Sub

