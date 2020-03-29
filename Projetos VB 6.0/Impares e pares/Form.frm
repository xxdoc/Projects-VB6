VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Impares e pares"
   ClientHeight    =   4185
   ClientLeft      =   8505
   ClientTop       =   3510
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   7305
   Begin VB.CommandButton cmdPar 
      Caption         =   "Pares"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdImp 
      Caption         =   "Impares"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.ListBox lstNum 
      Height          =   2400
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtN2 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtN1 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   $"Form.frx":0000
      Height          =   1695
      Left            =   2880
      TabIndex        =   7
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label lblNum2 
      Caption         =   "Número 2:"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblNum1 
      Caption         =   "Número 1:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Num1, Num2 As Integer
Private Sub cmdImp_Click()
lstNum.Clear
If Modex(txtN1) = True And Modex(txtN2) = True Then
    Num1 = txtN1.Text
    Num2 = txtN2.Text
    
    Do Until (Num1 = Num2 - 1)
        Num1 = Num1 + 1
        If Parimpa(Num1) = False Then
            lstNum.AddItem (Num1)
        End If
    Loop
Else
    MsgBox ("Insira números válidos")
End If
End Sub

Private Sub cmdPar_Click()
lstNum.Clear
If Modex(txtN1) = True And Modex(txtN2) = True Then
    Num1 = txtN1.Text
    Num2 = txtN2.Text
    
    Do Until (Num1 = Num2)
        Num1 = Num1 + 1
        If Parimpa(Num1) = True Then
            lstNum.AddItem (Num1)
        End If
    Loop
Else
    MsgBox ("Insira números válidos")
End If
End Sub

