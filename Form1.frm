VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   Caption         =   " Designed By ZZY"
   ClientHeight    =   9930
   ClientLeft      =   2460
   ClientTop       =   3645
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command10 
      Caption         =   "������ɫ?"
      BeginProperty Font 
         Name            =   "����"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14400
      TabIndex        =   21
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "ʱ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14400
      TabIndex        =   20
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "զ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14400
      TabIndex        =   19
      Top             =   1080
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Left            =   14760
      Top             =   2160
   End
   Begin VB.CommandButton Command7 
      Caption         =   "��APIָ��Ӣ��"
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "��APIָ������"
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "��ͣ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "���ѱ����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   10
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   9
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   3240
      Width           =   15015
   End
   Begin VB.Timer Timer1 
      Left            =   14760
      Top             =   2640
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "״ָ̬ʾ"
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "�������"
      Height          =   1095
      Left            =   4200
      TabIndex        =   14
      Top             =   2040
      Width           =   4575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Loading......"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   9600
      TabIndex        =   3
      Top             =   1080
      Width           =   3900
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Loading......"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      TabIndex        =   5
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label Label10 
      Caption         =   "New��"
      Height          =   255
      Left            =   13920
      TabIndex        =   22
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Build 05  2016.09.10 Designed by ZZY"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13440
      TabIndex        =   18
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Top             =   -120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ӣ�����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      TabIndex        =   6
      Top             =   -120
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2880
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "׼������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12000
      TabIndex        =   1
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "���ڽ���"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11280
      TabIndex        =   8
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "״̬��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      TabIndex        =   2
      Top             =   2280
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Timer1.Interval = 500
Label1.Caption = Year(Date) & "��" & Month(Date) & "��" & Day(Date) & "��"
'Label3.Caption = Time
a = 43
End Sub

Private Sub Command1_Click()
'begin
Label6.Visible = False
'����Ϊ����ʱ
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
'������ɫ
        Form1.Visible = False
        Form2.Visible = True
        Form3.Visible = False
        Form4.Visible = False
End Sub

Private Sub Command2_Click()
'finished
End
End Sub

Private Sub Command3_Click()
'save as
Open "C://Ӣ�����/201.txt" For Output As #1
Print #1, Text1.Text
Close #1
End Sub

Private Sub Command4_Click()
'open as
Open "C://Ӣ�����/today.txt" For Input As #1
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
MsgBox "����δ���ã�"
End Sub
Private Sub Timer1_Timer()
Label3.Caption = Time
End Sub
