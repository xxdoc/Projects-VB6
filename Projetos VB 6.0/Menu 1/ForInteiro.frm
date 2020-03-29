VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabuada Form Simples"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtNum 
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar"
      Height          =   855
      Left            =   2880
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
   End
   Begin VB.ListBox lstTabuada 
      Height          =   2985
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Joker As Integer
Private Sub cmdGerar_Click()
lstTabuada.Clear
If IsNumeric(txtNum.Text) Then
    For Joker = 1 To 10
        lstTabuada.AddItem (txtNum.Text & " x " & Format(Joker, "00") & " = " & Format(Joker * txtNum.Text, "000"))
    Next
End If
End Sub
