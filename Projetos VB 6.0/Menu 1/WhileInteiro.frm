VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "While"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7815
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7815
   Begin VB.ListBox lstTabuada 
      Height          =   2985
      Left            =   1080
      TabIndex        =   2
      Tag             =   "c"
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar"
      Height          =   2175
      Left            =   3600
      TabIndex        =   1
      Tag             =   "c"
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtNum 
      Height          =   615
      Left            =   3600
      TabIndex        =   0
      Tag             =   "c"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Line Line6 
      X1              =   960
      X2              =   960
      Y1              =   4440
      Y2              =   720
   End
   Begin VB.Line Line9 
      Tag             =   "c"
      X1              =   5760
      X2              =   3480
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line8 
      Tag             =   "c"
      X1              =   3480
      X2              =   5760
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line7 
      Tag             =   "c"
      X1              =   3480
      X2              =   960
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line5 
      Tag             =   "c"
      X1              =   960
      X2              =   5760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line4 
      Tag             =   "c"
      X1              =   5760
      X2              =   5760
      Y1              =   4440
      Y2              =   720
   End
   Begin VB.Line Line3 
      Tag             =   "c"
      X1              =   3480
      X2              =   5760
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line2 
      Tag             =   "c"
      X1              =   3480
      X2              =   3480
      Y1              =   720
      Y2              =   4440
   End
   Begin VB.Line Line1 
      Tag             =   "c"
      X1              =   960
      X2              =   3480
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label2 
      Caption         =   "Tabuada Gerada"
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Tag             =   "c"
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Digite a Tabuada!"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Tag             =   "c"
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Joker As Integer
Private Sub cmdGerar_Click()
lstTabuada.Clear
If IsNumeric(txtNum.Text) Then
    Joker = 1
    Do While Joker < 11
        lstTabuada.AddItem (txtNum.Text & " x " & Format(Joker, "00") & " = " & Format(Joker * txtNum.Text, "000"))
        Joker = Joker + 1
    Loop
End If
End Sub

