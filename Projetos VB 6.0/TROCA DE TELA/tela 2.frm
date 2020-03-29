VERSION 5.00
Begin VB.Form Form2 
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5775
   LinkTopic       =   "Form2"
   ScaleHeight     =   4275
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton volt 
      Caption         =   "Voltar"
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Title 
      Caption         =   "Mudar Título"
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Tela 2"
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
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Title_Click()
Form2.Caption = "Hoje é sexta!"
End Sub
Private Sub volt_Click()
Form1.Show
Form2.Hide
Form1.BackColor = vbButtonFace
End Sub
