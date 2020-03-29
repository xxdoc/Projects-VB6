VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd1 
      Caption         =   "Inverter"
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txt1 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lbl1 
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Palavra"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
lbl1.Caption = StrReverse(txt1.Text)
End Sub
