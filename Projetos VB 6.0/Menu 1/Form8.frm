VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form8"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7740
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7740
   Begin VB.TextBox txtQTD 
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2880
      Width           =   3375
   End
   Begin VB.CommandButton cmdNight 
      Caption         =   "Command1"
      Height          =   1335
      Left            =   5280
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtNight 
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label lblMid 
      Caption         =   "Label2"
      Height          =   975
      Left            =   1200
      TabIndex        =   4
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Right"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNight_Click()
lblRight.Caption = Right(txtNight.Text, txtQTD.Text)
End Sub
