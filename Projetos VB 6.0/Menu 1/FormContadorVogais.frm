VERSION 5.00
Begin VB.Form Form13 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form13"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7200
   Begin VB.TextBox txtPalavra 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton cmdSeparar 
      Caption         =   "Contar"
      Height          =   855
      Left            =   3960
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Contar Vogais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   4
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Digite uma algo"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblVogais 
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tia, Joker As Integer
Dim AAAA, OOOO As String
Dim EEEE, IIII As String
Dim UUUU, Skull As String
Dim Akechi, Batima As Integer
Private Sub cmdSeparar_Click()
Joker = 0
Akechi = Len(txtPalavra.Text)
Skull = LCase(txtPalavra.Text)
For Batima = 1 To Akechi
    Tia = InStr(Batima, Skull, AAAA)
    If Tia > 0 Then
        Joker = Joker + 1
        Batima = Tia
    End If
Next
For Batima = 1 To Akechi
    Tia = InStr(Batima, Skull, EEEE)
    If Tia > 0 Then
        Joker = Joker + 1
        Batima = Tia
    End If
Next
For Batima = 1 To Akechi
    Tia = InStr(Batima, Skull, IIII)
    If Tia > 0 Then
        Joker = Joker + 1
        Batima = Tia
    End If
Next
For Batima = 1 To Akechi
    Tia = InStr(Batima, Skull, OOOO)
    If Tia > 0 Then
        Joker = Joker + 1
        Batima = Tia
    End If
Next
For Batima = 1 To Akechi
    Tia = InStr(Batima, Skull, UUUU)
    If Tia > 0 Then
        Joker = Joker + 1
        Batima = Tia
    End If
Next
lblVogais.Caption = "Número de vogais = " & Joker
End Sub
Private Sub Form_Load()
AAAA = "a"
EEEE = "e"
IIII = "i"
OOOO = "o"
UUUU = "u"
End Sub
