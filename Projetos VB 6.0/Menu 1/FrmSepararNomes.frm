VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form11"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8940
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   8940
   Begin VB.TextBox txtNomeC 
      Height          =   405
      Left            =   2880
      TabIndex        =   1
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton cmdSeparar 
      Caption         =   "Processar"
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "           Separar Nomes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   960
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "Nome e Sobrenome"
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "           Nome"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "      Sobrenome"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblNome 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label lblSobrenome 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   3600
      Width           =   2415
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Joker As Integer
Private Sub cmdSeparar_Click()
Joker = InStr(txtNomeC.Text, " ")
lblNome.Caption = Left(txtNomeC.Text, Joker)
Joker = Len(txtNomeC.Text) - Joker
lblSobrenome.Caption = Right(txtNomeC.Text, Joker)
End Sub

