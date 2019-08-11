VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   7065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14895
   LinkTopic       =   "Form6"
   ScaleHeight     =   7065
   ScaleWidth      =   14895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "华文细黑"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   12840
      TabIndex        =   8
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "   Exit"
      BeginProperty Font 
         Name            =   "华文细黑"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   12840
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "授权版本                 Build 4.2.x            Designed By ZZY and Desperate"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   11160
      TabIndex        =   6
      Top             =   6240
      Width           =   3495
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " 自由早读"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   6360
      TabIndex        =   5
      Top             =   4560
      Width           =   3375
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " 化学"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   3720
      TabIndex        =   4
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " 生物"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   1080
      TabIndex        =   3
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " 英语"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   3720
      TabIndex        =   2
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " 语文"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   1080
      TabIndex        =   1
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "早读软件"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   4335
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   7095
      Left            =   0
      Picture         =   "Form6.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14895
   End
End
Attribute VB_Name = "Form6"
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

Option Explicit
'窗口透明常数
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'窗口透明API
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2
Const LWA_COLORKEY = &H1
Private Sub Form_Load()
Dim Rec As RECT, hRgn As Long
    GetWindowRect Me.hWnd, Rec
    hRgn = CreateRoundRectRgn(0, 0, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, 50, 50) '这里的两个50是分别用来设置圆角的宽和高的
    SetWindowRgn Me.hWnd, hRgn, True
'Randomize
'R = Int(Rnd * 256) '红色
'G = Int(Rnd * 256) '绿色
'B = Int(Rnd * 256) '蓝色
'Me.BackColor = rgRGB(R, G, B)
Dim rtn As Long
Me.BackColor = RGB(256, 256, 256) '设置一下窗口的颜色
rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hWnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hWnd, RGB(0, 0, 0), 150, LWA_COLORKEY
'SetLayeredWindowAttributes hwnd, RGB(256, 256, 256), 150, LWA_ALPHA
'RGB(0, 0, 0)参数就是要透明掉的颜色
'End Sub 'Private Sub Command1_Click()
'语文访客
'MsgBox "感谢使用"
'Form7.Visible = True
End Sub


'Private Sub Command2_Click()
'登录英语
'If Text1.Text = "123456" Then
'Form5.Visible = True
'Else
'MsgBox "请输入正确的口令"
'End If
'End Sub



'Private Sub Command2_Click()
'If Combo1.Text = "语文" Then
'Form6.Visible = False
'Form7.Visible = True
'End If
'If Combo1.Text = "生物" Then
'Form6.Visible = False
'Form9.Visible = True
'End If
'If Combo1.Text = "英语" Then
'Form6.Visible = False
'Form5.Visible = True
'End If
'If Combo1.Text = "自由早读" Then
'Form6.Visible = False
'Form10.Visible = True
'End If
'If Combo1.Text = "化学" Then
'Form6.Visible = False
'Form11.Visible = True
'End If
'End Sub

Private Sub Label2_Click()
Form6.Visible = False
Form14.Visible = True
Form7.Visible = True
End Sub
Private Sub Label3_Click()
Form6.Visible = False
Form12.Visible = True
Form5.Visible = True
End Sub

Private Sub Label4_Click()
Form6.Visible = False
Form15.Visible = True
Form9.Visible = True
End Sub

Private Sub Label5_Click()
Form6.Visible = False
Form16.Visible = True
Form11.Visible = True
End Sub

Private Sub Label6_Click()
Form6.Visible = False
Form17.Visible = True
Form10.Visible = True
End Sub

Private Sub Label8_Click()
End
End Sub

Private Sub Label9_Click()
Form6.Visible = False
Form13.Visible = True
End Sub
