VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "զ������"
   ClientHeight    =   1380
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton OKButton 
      Caption         =   "������"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "��Ȩ�汾                 Build 4.2.x            Designed By ZZY and Desperate"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "����ʽ������������Աȡ����ϵ��"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub OKButton_Click()
Me.Visible = False
End Sub
