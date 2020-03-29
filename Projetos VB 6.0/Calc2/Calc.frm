VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDiv 
      Caption         =   "Dividir"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdMul 
      Caption         =   "Multiplicar"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "Subtrair"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdSom 
      Caption         =   "Somar"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtN2 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtN1 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblR 
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblN2 
      Caption         =   "Número 2"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblN1 
      Caption         =   "Número 1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDiv_Click()
If IsNumeric(txtN1.Text) Then
    If IsNumeric(txtN2.Text) Then
        If CCur(txtN2.Text) = 0 Then
            Call MsgBox("Divisão por zero não é possível")
        Else
        lblR.Caption = CCur(txtN1.Text) / CCur(txtN2.Text)
        End If
    Else
        Call MsgBox("Um dos valores não é numerico")
    End If
Else
    Call MsgBox("Um dos valores não é numerico")
End If
End Sub

Private Sub cmdMul_Click()
If IsNumeric(txtN1.Text) Then
    If IsNumeric(txtN2.Text) Then
        lblR.Caption = CCur(txtN1.Text) * CCur(txtN2.Text)
    Else
        Call MsgBox("Um dos valores não é numerico")
    End If
Else
    Call MsgBox("Um dos valores não é numerico")
End If
End Sub


Private Sub cmdSom_Click()
If IsNumeric(txtN1.Text) Then
    If IsNumeric(txtN2.Text) Then
        lblR.Caption = CCur(txtN1.Text) + CCur(txtN2.Text)
    Else
        Call MsgBox("Um dos valores não é numerico")
    End If
Else
    Call MsgBox("Um dos valores não é numerico")
End If
End Sub

Private Sub cmdSub_Click()
If IsNumeric(txtN1.Text) Then
    If IsNumeric(txtN2.Text) Then
        lblR.Caption = CCur(txtN1.Text) - CCur(txtN2.Text)
    Else
        Call MsgBox("Um dos valores não é numerico")
    End If
Else
    Call MsgBox("Um dos valores não é numerico")
End If
End Sub

Private Sub Form_Load()

End Sub
