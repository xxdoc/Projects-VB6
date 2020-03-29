VERSION 5.00
Begin VB.Form Primeira 
   Caption         =   "Primeira Prova de IDA"
   ClientHeight    =   4650
   ClientLeft      =   7365
   ClientTop       =   3300
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   7770
   Begin VB.CommandButton cmdMSG 
      Caption         =   "Exibir Mensagem"
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox txtEnderecoDaPessoa 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   5775
   End
   Begin VB.TextBox txtNomeDaPessoa 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label lblEnderecoDaPessoa 
      Caption         =   "Endereço"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblNomeDaPessoa 
      Caption         =   "Nome"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Primeira"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nome As Variant
Dim Endereco As Variant
Private Sub cmdMSG_Click()
Call MsgBox(Nome)
Call MsgBox(Endereco)
End Sub
Private Sub txtEnderecoDaPessoa_Change()
Endereco = txtEnderecoDaPessoa.Text
End Sub
Private Sub txtNomeDaPessoa_Change()
Nome = txtNomeDaPessoa.Text
End Sub
