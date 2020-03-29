VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form6"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5580
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5580
   Begin VB.CommandButton cmdLer 
      Caption         =   "Processar"
      Height          =   615
      Left            =   3840
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtLer 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Line Line5 
      X1              =   3840
      X2              =   3840
      Y1              =   3000
      Y2              =   2280
   End
   Begin VB.Line Line4 
      X1              =   3840
      X2              =   720
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line3 
      X1              =   720
      X2              =   720
      Y1              =   2280
      Y2              =   3000
   End
   Begin VB.Line Line2 
      X1              =   2280
      X2              =   2280
      Y1              =   2280
      Y2              =   3000
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   3840
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lblResposta 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   2400
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Resposta"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Digite algo"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Função Len"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdLer_Click()
lblResposta.Caption = Len(txtLer.Text)
End Sub

