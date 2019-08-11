VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   10185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15300
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10185
   ScaleWidth      =   15300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Left            =   14640
      Top             =   2760
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Text            =   "Form3.frx":0000
      Top             =   3360
      Width           =   15015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "结束"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   9
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "另存早读任务"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   8
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "打开已保存的"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   7
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "暂停"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "将API指向语文"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "将API指向英语"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Left            =   14640
      Top             =   2280
   End
   Begin VB.CommandButton Command8 
      Caption         =   "咋用来着"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14400
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "时差"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14400
      TabIndex        =   1
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "换个颜色?"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14400
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "任务操作"
      Height          =   1095
      Left            =   4200
      TabIndex        =   12
      Top             =   2160
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "状态指示"
      Height          =   1095
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "Loading......"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   9480
      TabIndex        =   17
      Top             =   1200
      Width           =   3900
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "准备进行"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12000
      TabIndex        =   20
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "现在是"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2760
      TabIndex        =   19
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Loading......"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      TabIndex        =   18
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "语文"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Build 05  2016.09.10 Designed by ZZY"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13440
      TabIndex        =   14
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "New！"
      Height          =   255
      Left            =   13800
      TabIndex        =   13
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "英语早读"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      TabIndex        =   16
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "正在进行"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11280
      TabIndex        =   21
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "状态："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      TabIndex        =   22
      Top             =   2400
      Width           =   2175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Interval = 500
Label1.Caption = Year(Date) & "年" & Month(Date) & "月" & Day(Date) & "日"
'Label3.Caption = Time
a = 43
End Sub

Private Sub Command1_Click()
'begin
Label6.Visible = False
'以下为倒计时
'Timer2.Interval = 1000
'Timer2.Enabled = True
'If a < 60 Then
'm = a
'Else
'h = s \ 60
'm = a Mod 60
'End If
'm = m - 1
's = 60
End Sub

Private Sub Command10_Click()
'换个颜色
        Form3.Visible = False
        Form2.Visible = True
        Form1.Visible = False
        Form4.Visible = False
End Sub

Private Sub Command2_Click()
'finished
End
End Sub

Private Sub Command3_Click()
'save as
Open "C://英语早读/201.txt" For Output As #1
Print #1, Text1.Text
Close #1
End Sub

Private Sub Command4_Click()
'open as
Open "C://英语早读/today.txt" For Input As #1
Do Until EOF(1)
Line Input #1, nextline
Text1.Text = Text1.Text + nextline + Chr(13) + Chr(10)
Loop
Close #1
End Sub

Private Sub Command5_Click()
'pause
Label6.Visible = True
End Sub

Private Sub Command6_Click()
'API to yuwen
Label8.Visible = True
        Command7.Visible = True
        Command6.Visible = False

End Sub

Private Sub Command7_Click()
'api to yingyu
Label8.Visible = False
        Command6.Visible = True
        Command7.Visible = False

End Sub

Private Sub Command8_Click()
Dialog.Visible = True
End Sub

Private Sub Command9_Click()
MsgBox "功能未启用！"
End Sub
Private Sub Timer1_Timer()
Label3.Caption = Time
End Sub

