VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form9"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7515
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7515
   Begin VB.TextBox txtIni 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Processar"
      Height          =   735
      Left            =   5040
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txt2 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   1080
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox txt1 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblexibir 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      Height          =   855
      Left            =   1320
      TabIndex        =   6
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Texto2"
      BeginProperty Font 
         Name            =   "Niagara Engraved"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Texto1"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "InStr"
      BeginProperty Font 
         Name            =   "MS Mincho"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
lblexibir.Caption = InStr(txtIni.Text, txt1.Text, txt2.Text)
End Sub

