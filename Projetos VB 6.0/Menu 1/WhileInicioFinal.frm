VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form4"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10845
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   10845
   Begin VB.TextBox txtNum2 
      Height          =   615
      Left            =   7080
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtNum 
      Height          =   615
      Left            =   4800
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar"
      Height          =   2175
      Left            =   4800
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin VB.ListBox lstTabuada 
      Height          =   2985
      Left            =   2280
      TabIndex        =   0
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Digite a Tabuada final!"
      Height          =   255
      Left            =   7080
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.Line Line10 
      X1              =   9000
      X2              =   9000
      Y1              =   720
      Y2              =   2040
   End
   Begin VB.Label Label1 
      Caption         =   "Digite a Tabuada inicial!"
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Tabuada Gerada"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   2160
      X2              =   4680
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      X1              =   4680
      X2              =   4680
      Y1              =   720
      Y2              =   4440
   End
   Begin VB.Line Line3 
      X1              =   4680
      X2              =   6960
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line4 
      X1              =   6960
      X2              =   6960
      Y1              =   4440
      Y2              =   720
   End
   Begin VB.Line Line5 
      X1              =   2160
      X2              =   9000
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line7 
      X1              =   4680
      X2              =   2160
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line8 
      X1              =   4680
      X2              =   9000
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line9 
      X1              =   9000
      X2              =   4680
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line6 
      X1              =   2160
      X2              =   2160
      Y1              =   4440
      Y2              =   720
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Igor As Integer
Dim Joker, Akechi As Integer
Private Sub cmdGerar_Click()
lstTabuada.Clear
If IsNumeric(txtNum.Text) Then
    Joker = 1
    Igor = txtNum.Text
    Akechi = txtNum2.Text - txtNum.Text
    Do While Akechi <> 0
        If Joker = 11 Then
            Joker = 1
            Akechi = Akechi - 1
            Igor = Igor + 1
            lstTabuada.AddItem ("--------------------------------------------")
        End If
        Do While Joker < 11
            lstTabuada.AddItem (Igor & " x " & Format(Joker, "00") & " = " & Format(Joker * Igor, "000"))
            Joker = Joker + 1
        Loop
    Loop
End If
End Sub


