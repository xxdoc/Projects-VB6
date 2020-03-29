VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEndereco 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtNome 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton cmdMensagem 
      Caption         =   "Exibir mensagem"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblEndereco 
      Caption         =   "Endereço"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblNome 
      Caption         =   "Nome"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMensagem_Click()
MsgBox txtNome.Text & " mora em " & txtEndereco.Text
End Sub

