VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   4245
   ClientLeft      =   10305
   ClientTop       =   1605
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6045
   Begin VB.CommandButton Avc 
      Caption         =   "Avançar"
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cor 
      Caption         =   "Mudar Cor"
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Tela 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Avc_Click()
Form1.Hide
Form2.Show
Label1.BackColor = vbButtonFace
End Sub
Private Sub cor_Click()
Form1.BackColor = vbWhite
Label1.BackColor = vbWhite
End Sub

