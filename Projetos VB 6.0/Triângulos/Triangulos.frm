VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detector de Triangulos"
   ClientHeight    =   1875
   ClientLeft      =   9705
   ClientTop       =   4470
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   3375
   Begin VB.CommandButton cmdVerificar 
      Caption         =   "Verificar "
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox txtLadoC 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtLadoB 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtLadoA 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblLadoC 
      Caption         =   "Lado C"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblLadoB 
      Caption         =   "Lado B"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblLadoA 
      Caption         =   "Lado A"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Feito por Pedro Henrique'
'R.A.:2190063-7'
Dim LadoC As String
Dim LadoB As String
Dim LadoA As String
Private Sub cmdVerificar_Click()
LadoA = txtLadoA.Text
LadoB = txtLadoB.Text
LadoC = txtLadoC.Text
If CCur(LadoA) < 100000 Or CCur(LadoB) < 100000 Or CCur(LadoC) < 100000 Then
    If IsNumeric(LadoA) = True And IsNumeric(LadoB) = True And IsNumeric(LadoC) = True Then
        If CCur(LadoA) < 0 Or CCur(LadoB) < 0 Or CCur(LadoC) < 0 Then
            Call MsgBox("Insira números positivos")
        End If
        If CCur(LadoA) = 0 Or CCur(LadoB) = 0 Or CCur(LadoC) = 0 Then
            Call MsgBox("Insira números maiores que zero")
        End If
        If CCur(LadoB) - CCur(LadoC) < CCur(LadoA) < CCur(LadoB) + CCur(LadoC) And CCur(LadoA) - CCur(LadoC) < CCur(LadoB) < CCur(LadoA) + CCur(LadoC) And CCur(LadoA) - CCur(LadoB) < CCur(LadoC) < CCur(LadoA) + CCur(LadoB) Then
            If CCur(LadoA) = CCur(LadoC) And CCur(LadoA) = CCur(LadoB) Then
                Call MsgBox("Os valores indicam um Triangulo Equilátero")
            End If
            If (CCur(LadoA) = CCur(LadoB) And CCur(LadoA) <> CCur(LadoC)) Or (CCur(LadoA) = CCur(LadoC) And CCur(LadoA) <> CCur(LadoB)) Or (CCur(LadoB) = CCur(LadoC) And CCur(LadoB) <> CCur(LadoA)) Then
                Call MsgBox("Os valores indicam um Triangulo Isósceles")
            End If
            If CCur(LadoA) <> CCur(LadoB) And CCur(LadoA) <> CCur(LadoC) And CCur(LadoC) <> CCur(LadoB) Then
                Call MsgBox("Os valores indicam um Triangulo Escaleno")
            End If
        Else
            Call MsgBox("Os valores indicados não pertencem à um triangulo")
        End If
    Else
        If LadoA = "" Or LadoB = "" Or LadoC = "" Then
            Call MsgBox("Insira todos os valores")
        Else
            Call MsgBox("Insira apenas números")
        End If
    End If
Else
    Call MsgBox("Insira números menores")
End If
End Sub


