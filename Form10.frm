VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form10"
   ClientHeight    =   10005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15240
   LinkTopic       =   "Form10"
   ScaleHeight     =   10005
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "���׹رմ˳���"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tool Box"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Form10.frx":0000
      Top             =   1920
      Width           =   14895
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "How..."
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   120
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   14400
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading......"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   10800
      TabIndex        =   11
      Top             =   0
      Width           =   7395
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�������:"
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
      Left            =   1800
      TabIndex        =   10
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Control of time."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   9
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
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
      Height          =   1095
      Index           =   0
      Left            =   6960
      TabIndex        =   8
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6240
      TabIndex        =   7
      Top             =   720
      Width           =   3735
   End
End
Attribute VB_Name = "Form10"
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
'����͸������
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'����͸��API
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H1
Const LWA_COLORKEY = &H2

Private Sub Form_Load()
Timer1.Interval = 500
'Label1.Caption = Year(Date) & "��" & Month(Date) & "��" & Day(Date) & "��"
'Label3.Caption = Time
'a = 43
Dim Rec As RECT, hRgn As Long
    GetWindowRect Me.hWnd, Rec
    hRgn = CreateRoundRectRgn(0, 0, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, 50, 50) '���������50�Ƿֱ���������Բ�ǵĿ�͸ߵ�
    SetWindowRgn Me.hWnd, hRgn, True
'Randomize
'R = Int(Rnd * 256) '��ɫ
'G = Int(Rnd * 256) '��ɫ
'B = Int(Rnd * 256) '��ɫ
'Me.BackColor = rgRGB(R, G, B)
Dim rtn As Long
Me.BackColor = RGB(256, 256, 256) '����һ�´��ڵ���ɫ
rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hWnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hWnd, RGB(256, 256, 256), 210, LWA_COLORKEY
'RGB(0, 0, 0)��������Ҫ͸��������ɫ
Timer1.Interval = 500
'Label1.Caption = Year(Date) & "��" & Month(Date) & "��" & Day(Date) & "��"
'Label3.Caption = Time
'a = 43
End Sub
Private Sub Command2_Click()
'START,2
Label6(0).Visible = False
Command2.Visible = False
Command3.Visible = True
End Sub

Private Sub Command3_Click()
Label6(0).Visible = True
Command2.Visible = True
Command3.Visible = False
End Sub

'Private Sub Command3_Click()
'PAUSE,2
'Label6.Visible = False
'End Sub

Private Sub Command4_Click()
'TOOLBOX,2
Form8.Visible = True
End Sub

'Private Sub Command1_Click()
'begin
'Label6.Visible = False
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
'End Sub

'Private Sub Command10_Click()
'������ɫ
        'Form3.Visible = False
        'Form2.Visible = True
        'Form1.Visible = False
        'Form4.Visible = False
'End Sub

'Private Sub Command2_Click()
'finished
'End
'End Sub

'Private Sub Command3_Click()
'save as
'Open "C://Ӣ�����/201.txt" For Output As #1
'Print #1, Text1.Text
'Close #1
'End Sub

'Private Sub Command4_Click()
'open as
'Open "C://Ӣ�����/today.txt" For Input As #1
'Do Until EOF(1)
'Line Input #1, nextline
'Text1.Text = Text1.Text + nextline + Chr(13) + Chr(10)
'Loop
'Close #1
'End Sub

Private Sub Command5_Click()
'NEW,2
'Form2.Visible = True
MsgBox "����δ���ã��������Աȡ����ϵ��"
End Sub

Private Sub Command1_Click()
End
End Sub

'Private Sub Command6_Click()
'API to yuwen
'Label1.Visible = True
       ' Command7.Visible = True
       ' Command6.Visible = False
'End Sub

'Private Sub Command7_Click()
'api TO YINGYU
'Label1.Visible = False
      '  Command6.Visible = True
      '  Command7.Visible = False
'End Sub

Private Sub Command8_Click()
'ABOUT,2
Dialog.Visible = True
End Sub

'Private Sub Command8_Click()
'Dialog.Visible = True
'End Sub

'Private Sub Command9_Click()
'MsgBox "����δ���ã�"
'End Sub



Private Sub Timer1_Timer()
Label3.Caption = Time
End Sub



