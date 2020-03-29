VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form7"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6285
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6285
   Begin VB.TextBox txtQTD 
      Height          =   405
      Left            =   2400
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtLer 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdLer 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Processar"
      Height          =   615
      Left            =   3720
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "QTD"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Tag             =   "QTD"
      Top             =   1320
      Width           =   615
   End
   Begin VB.Line Line3 
      X1              =   600
      X2              =   600
      Y1              =   2400
      Y2              =   3120
   End
   Begin VB.Line Line1 
      X1              =   3720
      X2              =   600
      Y1              =   3120
      Y2              =   3120
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
      Left            =   960
      TabIndex        =   5
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Digite algo"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Resposta"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblResposta 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   2160
      X2              =   2160
      Y1              =   2400
      Y2              =   3120
   End
   Begin VB.Line Line4 
      X1              =   3720
      X2              =   600
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line5 
      X1              =   3720
      X2              =   3720
      Y1              =   3120
      Y2              =   2400
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLer_Click()
lblResposta.Caption = Left(txtLer.Text, txtQTD.Text)
End Sub
