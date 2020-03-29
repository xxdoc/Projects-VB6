VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lsbTabuada 
      Height          =   4155
      ItemData        =   "Tabuada.frx":0000
      Left            =   600
      List            =   "Tabuada.frx":0002
      TabIndex        =   3
      Top             =   1320
      Width           =   3135
   End
   Begin VB.CommandButton cmdtab 
      Caption         =   "Calcular tabuada"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox txtTab 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Número"
      Height          =   255
      Left            =   600
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
Dim Nume1, Nume3 As Currency
Private Sub cmdtab_Click()
lsbTabuada.Clear
Nume1 = txtTab
    If IsNumeric(Nume1) = True Then
        Nume2 = 0
        Dim Contador As Integer
        For Contador = 1 To 20
            Nume3 = Nume1 * Contador
            lsbTabuada.AddItem (Nume1 & " x " & Format(Contador, "00") & " = " & Format(Nume3, "000"))
        Next Contador
    Else
        MsgBox ("Insira um número")
    End If
End Sub
