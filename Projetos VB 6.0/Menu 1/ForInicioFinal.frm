VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form5"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9045
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   9045
   Tag             =   "-----------"
   Begin VB.ListBox lstTabuada 
      Height          =   2985
      Left            =   1200
      TabIndex        =   3
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar"
      Height          =   2175
      Left            =   3720
      TabIndex        =   2
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtNum 
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtNum2 
      Height          =   615
      Left            =   6000
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Line Line6 
      X1              =   1080
      X2              =   1080
      Y1              =   4680
      Y2              =   960
   End
   Begin VB.Line Line9 
      X1              =   7920
      X2              =   3600
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line8 
      X1              =   3600
      X2              =   7920
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line7 
      X1              =   3600
      X2              =   1080
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line5 
      X1              =   1080
      X2              =   7920
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line4 
      X1              =   5880
      X2              =   5880
      Y1              =   4680
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   3600
      X2              =   5880
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line2 
      X1              =   3600
      X2              =   3600
      Y1              =   960
      Y2              =   4680
   End
   Begin VB.Line Line1 
      X1              =   1080
      X2              =   3600
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label2 
      Caption         =   "Tabuada Gerada"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Digite a Tabuada inicial!"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Line Line10 
      X1              =   7920
      X2              =   7920
      Y1              =   960
      Y2              =   2280
   End
   Begin VB.Label Label3 
      Caption         =   "Digite a Tabuada final!"
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "Form5"
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
    For Akechi = txtNum.Text - 1 To txtNum2.Text - 1
        For Joker = 1 To 10
            lstTabuada.AddItem (Igor & " x " & Format(Joker, "00") & " = " & Format(Joker * Igor, "000"))
        Next
        Igor = Igor + 1
        lstTabuada.AddItem ("--------------------------------------------")
    Next
End If
End Sub



