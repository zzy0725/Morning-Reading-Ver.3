VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Color"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   1920
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   2250
   ScaleWidth      =   1920
   StartUpPosition =   2  '屏幕中心
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "淡雅米黄"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "清新天蓝"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "护眼青绿"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0FFFF&
      BorderColor     =   &H00C0FFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   120
      Top             =   1560
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFC0&
      BorderColor     =   &H00FFFFC0&
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   120
      Top             =   840
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderColor     =   &H00C0FFC0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
'GREEN
        Form2.Visible = False
        Form1.Visible = True
        Form3.Visible = False
        Form4.Visible = False
End Sub

Private Sub Label2_Click()
'BULE
        Form2.Visible = False
        Form3.Visible = True
        Form1.Visible = False
        Form4.Visible = False
End Sub

Private Sub Label3_Click()
'YELLOW
        Form2.Visible = False
        Form4.Visible = True
        Form3.Visible = False
        Form1.Visible = False
End Sub
