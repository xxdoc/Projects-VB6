VERSION 5.00
Begin VB.Form Form12 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form12"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6855
   Begin VB.CommandButton cmdSeparar 
      Caption         =   "Separar"
      Height          =   855
      Left            =   4080
      TabIndex        =   1
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtPalavra 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblVogais 
      Height          =   495
      Left            =   2040
      TabIndex        =   15
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Contador de Vogais"
      Height          =   375
      Left            =   480
      TabIndex        =   14
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Digite uma palavra"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Separar Vogais"
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
      Left            =   3120
      TabIndex        =   12
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "A -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "E -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblE 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "I -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "O -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label lblI 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label lblO 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "U -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label lblU 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   4080
      Width           =   495
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tia, Joker As Integer
Dim AAAA, OOOO As String
Dim EEEE, IIII As String
Dim UUUU, Skull As String
Dim Akechi, Batima As Integer
Dim China, Japao As Integer
Dim NKorea, Skorea As Integer
Private Sub cmdSeparar_Click()
Joker = 0
China = 0
Japao = 0
Skorea = 0
NKorea = 0
Akechi = Len(txtPalavra.Text)
Skull = LCase(txtPalavra.Text)
For Batima = 1 To Akechi
    Tia = InStr(Batima, Skull, AAAA)
    If Tia > 0 Then
        Joker = Joker + 1
        Batima = Tia
    End If
Next
lblA.Caption = Joker
For Batima = 1 To Akechi
    Tia = InStr(Batima, Skull, EEEE)
    If Tia > 0 Then
        Joker = Joker + 1
        Batima = Tia
        China = China + 1
    End If
Next
lblE.Caption = China
For Batima = 1 To Akechi
    Tia = InStr(Batima, Skull, IIII)
    If Tia > 0 Then
        Joker = Joker + 1
        Batima = Tia
        Japao = Japao + 1
    End If
Next
lblI.Caption = Japao
For Batima = 1 To Akechi
    Tia = InStr(Batima, Skull, OOOO)
    If Tia > 0 Then
        Joker = Joker + 1
        Batima = Tia
        NKorea = NKorea + 1
    End If
Next
lblO.Caption = NKorea
For Batima = 1 To Akechi
    Tia = InStr(Batima, Skull, UUUU)
    If Tia > 0 Then
        Joker = Joker + 1
        Batima = Tia
        Skorea = Skorea + 1
    End If
Next
lblU.Caption = Skorea
lblVogais.Caption = Joker
End Sub

Private Sub Form_Load()
AAAA = "a"
EEEE = "e"
IIII = "i"
OOOO = "o"
Joker = 0
China = 0
Japao = 0
Skorea = 0
NKorea = 0
UUUU = "u"
End Sub

