VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form8"
   ClientHeight    =   3900
   ClientLeft      =   825
   ClientTop       =   1005
   ClientWidth     =   3900
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form8"
   ScaleHeight     =   3900
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "关闭计算机"
      Default         =   -1  'True
      Height          =   1095
      Left            =   1320
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "结束explorer"
      Height          =   1095
      Left            =   1320
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "旧版API入口"
      Height          =   1095
      Left            =   2520
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "返回早读界面"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BorderColor     =   &H00404040&
      DrawMode        =   3  'Not Merge Pen
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   3735
      Left            =   0
      Shape           =   5  'Rounded Square
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Command1_Click()
Shell "taskkill /im explorer.exe /f"
End Sub

Private Sub Command4_Click()
Shell "shutdown /s"
End Sub

Private Sub Form_Load()
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
SetLayeredWindowAttributes hWnd, RGB(256, 256, 256), 20, LWA_COLORKEY
'RGB(0, 0, 0)参数就是要透明掉的颜色
End Sub

'Private Sub Command1_Click()
'打开保存的
'Open "C://英语早读/today.txt" For Input As #1
'Do Until EOF(1)
'Line Input #1, nextline
'Text1.Text = Text1.Text + nextline + Chr(13) + Chr(10)
'Loop
'Close #1
'End Sub

Private Sub Command2_Click()
Form8.Visible = False
End Sub

Private Sub Command3_Click()
'旧版API
Form2.Visible = True
Form8.Visible = False
Form5.Visible = False
End Sub

'Private Sub Command4_Click()
'另存早读任务
'Open "C://英语早读/201.txt" For Output As #1
'Print #1, Text1.Text
'Close #1
'End Sub

