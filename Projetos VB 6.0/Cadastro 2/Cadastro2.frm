VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   1320
      Width           =   2055
      Begin VB.OptionButton Option1 
         Caption         =   "Masculino"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Feminino"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   14
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   1680
      Width           =   2655
      Begin VB.OptionButton Option4 
         Caption         =   "Noite"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   12
         Top             =   0
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Tarde"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   1
         Top             =   0
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Manhã"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cadastrar"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Cadastro2.frx":0000
      Left            =   720
      List            =   "Cadastro2.frx":000D
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "R.A."
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Período"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Sexo"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4680
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label3 
      Caption         =   "Curso"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Aluno"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call MsgBox("Cadastro Efetuado")
End Sub
