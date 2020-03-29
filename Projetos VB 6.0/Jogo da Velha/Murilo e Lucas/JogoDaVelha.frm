VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Jogo Da Velha"
   ClientHeight    =   4428
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   6444
   DrawMode        =   16  'Merge Pen
   BeginProperty Font 
      Name            =   "Bookman Old Style"
      Size            =   7.8
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4428
   ScaleWidth      =   6444
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptionO 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4080
      TabIndex        =   12
      Top             =   2880
      Width           =   1092
   End
   Begin VB.OptionButton OptionX 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2760
      TabIndex        =   11
      Top             =   2880
      Width           =   1092
   End
   Begin VB.CommandButton cmdZP 
      Caption         =   "Zerar Placar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   360
      TabIndex        =   9
      Top             =   3240
      Width           =   1332
   End
   Begin VB.CommandButton cmdNJ 
      Caption         =   "Novo Jogo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   360
      TabIndex        =   8
      Top             =   2520
      Width           =   1332
   End
   Begin VB.Line Line16 
      X1              =   5880
      X2              =   2880
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line15 
      X1              =   2880
      X2              =   2880
      Y1              =   600
      Y2              =   2400
   End
   Begin VB.Line Line14 
      X1              =   2880
      X2              =   5880
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line13 
      X1              =   2880
      X2              =   5880
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line12 
      X1              =   5880
      X2              =   5880
      Y1              =   600
      Y2              =   2400
   End
   Begin VB.Line Line11 
      X1              =   2880
      X2              =   5880
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line10 
      X1              =   3840
      X2              =   3840
      Y1              =   600
      Y2              =   2400
   End
   Begin VB.Line Line9 
      X1              =   3960
      X2              =   3960
      Y1              =   2760
      Y2              =   3480
   End
   Begin VB.Line Line8 
      X1              =   5280
      X2              =   5280
      Y1              =   3480
      Y2              =   2760
   End
   Begin VB.Line Line7 
      X1              =   2640
      X2              =   5280
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line6 
      X1              =   5280
      X2              =   2640
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line5 
      X1              =   2640
      X2              =   2640
      Y1              =   2760
      Y2              =   3480
   End
   Begin VB.Line LineDD 
      Visible         =   0   'False
      X1              =   360
      X2              =   2160
      Y1              =   2280
      Y2              =   120
   End
   Begin VB.Line LineVD 
      Visible         =   0   'False
      X1              =   1920
      X2              =   1920
      Y1              =   240
      Y2              =   2160
   End
   Begin VB.Label lblRV 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3960
      TabIndex        =   19
      Top             =   1920
      Width           =   1812
   End
   Begin VB.Label lblRO 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3960
      TabIndex        =   18
      Top             =   1320
      Width           =   1812
   End
   Begin VB.Label lblRX 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3960
      TabIndex        =   17
      Top             =   720
      Width           =   1812
   End
   Begin VB.Label Label4 
      Caption         =   "   X -"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3000
      TabIndex        =   16
      Top             =   720
      Width           =   732
   End
   Begin VB.Label Label3 
      Caption         =   "   O -"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3000
      TabIndex        =   15
      Top             =   1320
      Width           =   732
   End
   Begin VB.Label Label2 
      Caption         =   "Velha -"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3000
      TabIndex        =   14
      Top             =   1920
      Width           =   732
   End
   Begin VB.Line LineDE 
      Visible         =   0   'False
      X1              =   2160
      X2              =   360
      Y1              =   2280
      Y2              =   120
   End
   Begin VB.Line LineHS 
      Visible         =   0   'False
      X1              =   360
      X2              =   2280
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line LineVE 
      Visible         =   0   'False
      X1              =   600
      X2              =   600
      Y1              =   2160
      Y2              =   240
   End
   Begin VB.Label lblSE 
      Caption         =   "    "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   13
      Top             =   240
      Width           =   375
   End
   Begin VB.Line LineHC 
      Visible         =   0   'False
      X1              =   360
      X2              =   2280
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line LineHI 
      Visible         =   0   'False
      X1              =   360
      X2              =   2280
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line LineVC 
      Visible         =   0   'False
      X1              =   1320
      X2              =   1320
      Y1              =   2160
      Y2              =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Placar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3840
      TabIndex        =   10
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label lblID 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1680
      TabIndex        =   7
      Top             =   1680
      Width           =   492
   End
   Begin VB.Label lblIC 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1080
      TabIndex        =   6
      Top             =   1680
      Width           =   372
   End
   Begin VB.Label lblIE 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblCD 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1680
      TabIndex        =   4
      Top             =   960
      Width           =   492
   End
   Begin VB.Label lblCC 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1080
      TabIndex        =   3
      Top             =   960
      Width           =   372
   End
   Begin VB.Label lblCE 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblSD 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   492
   End
   Begin VB.Label lblSC 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   360
      X2              =   2160
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   360
      X2              =   2160
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   1560
      X2              =   1560
      Y1              =   240
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   960
      X2              =   960
      Y1              =   240
      Y2              =   2160
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Velha As Integer
Private Sub cmdNJ_Click()
Velha = 0
lblSE.Enabled = True
lblSC.Enabled = True
lblSD.Enabled = True
lblCE.Enabled = True
lblCC.Enabled = True
lblCD.Enabled = True
lblIE.Enabled = True
lblIC.Enabled = True
lblID.Enabled = True
lblSE.Caption = ""
lblSC.Caption = ""
lblSD.Caption = ""
lblCE.Caption = ""
lblCC.Caption = ""
lblCD.Caption = ""
lblIE.Caption = ""
lblIC.Caption = ""
lblID.Caption = ""
LineVC.Visible = False
LineHS.Visible = False
LineHC.Visible = False
LineHI.Visible = False
LineVD.Visible = False
LineVE.Visible = False
LineDD.Visible = False
LineDE.Visible = False
If OptionX.Enabled = True Then
    OptionX.Enabled = False
    OptionO.Enabled = True
    OptionX.Value = False
    OptionO.Value = True
    Else
    OptionX.Enabled = True
    OptionO.Enabled = False
    OptionX.Value = True
    OptionO.Value = False
End If
End Sub

Private Sub cmdZP_Click()
lblRX.Caption = "0"
lblRO.Caption = "0"
lblRV.Caption = "0"
End Sub

Private Sub lblCC_Click()
If OptionX.Value = True Then
    lblCC.Caption = "X"
    Velha = Velha + 1
    ElseIf OptionO.Value = True Then
    lblCC.Caption = "O"
    Velha = Velha + 1
End If
If lblCE.Caption = "X" And lblCC.Caption = "X" And lblCD.Caption = "X" Then
    LineHC.Visible = True
    Velha = 0
End If
If lblSD.Caption = "X" And lblCD.Caption = "X" And lblID.Caption = "X" Then
    LineVD.Visible = True
    Velha = 0
End If
If lblCC.Caption = "X" And lblSC.Caption = "X" And lblIC.Caption = "X" Then
    LineVC.Visible = True
    Velha = 0
End If
If lblIE.Caption = "X" And lblCC.Caption = "X" And lblSD.Caption = "X" Then
    LineDD.Visible = True
    Velha = 0
End If
If lblID.Caption = "X" And lblCC.Caption = "X" And lblSE.Caption = "X" Then
    LineDE.Visible = True
    Velha = 0
End If
If lblCE.Caption = "X" And lblCC.Caption = "X" And lblCD.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblSC.Caption = "X" And lblCC.Caption = "X" And lblIC.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblIE.Caption = "X" And lblCC.Caption = "X" And lblSD.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblID.Caption = "X" And lblCC.Caption = "X" And lblSE.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblCE.Caption = "O" And lblCC.Caption = "O" And lblCD.Caption = "O" Then
    LineHC.Visible = True
    Velha = 0
End If
If lblSD.Caption = "O" And lblCD.Caption = "O" And lblID.Caption = "O" Then
    LineVD.Visible = True
    Velha = 0
End If
If lblCC.Caption = "O" And lblSC.Caption = "O" And lblIC.Caption = "O" Then
    LineVC.Visible = True
    Velha = 0
End If
If lblIE.Caption = "O" And lblCC.Caption = "O" And lblSD.Caption = "O" Then
    LineDD.Visible = True
    Velha = 0
End If
If lblSC.Caption = "O" And lblCC.Caption = "O" And lblIC.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblCE.Caption = "O" And lblCC.Caption = "O" And lblCD.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblIE.Caption = "O" And lblCC.Caption = "O" And lblSD.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblID.Caption = "O" And lblCC.Caption = "O" And lblSE.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblID.Caption = "O" And lblCC.Caption = "O" And lblSE.Caption = "O" Then
    LineDE.Visible = True
    Velha = 0
End If
If lblCC.Caption = "X" Then
    OptionX.Enabled = False
    OptionO.Enabled = True
    OptionX.Value = False
    OptionO.Value = True
    Else
    OptionX.Enabled = True
    OptionO.Enabled = False
    OptionX.Value = True
    OptionO.Value = False
End If
If Velha = 9 Then
    lblRV.Caption = lblRV.Caption + 1
End If
If lblCC.Caption = "X" Or lblCC.Caption = "O" Then
    lblCC.Enabled = False
End If
If LineHS.Visible = True Or LineHC.Visible = True Or LineHI.Visible = True Or LineVE.Visible = True Or LineVC.Visible = True Or LineVD.Visible = True Or LineDE.Visible = True Or LineDD.Visible = True Then
    lblSE.Enabled = False
    lblSC.Enabled = False
    lblSD.Enabled = False
    lblIE.Enabled = False
    lblIC.Enabled = False
    lblID.Enabled = False
    lblCE.Enabled = False
    lblCC.Enabled = False
    lblCD.Enabled = False
End If
End Sub

Private Sub lblCD_Click()
If OptionX.Value = True Then
    lblCD.Caption = "X"
    Velha = Velha + 1
    ElseIf OptionO.Value = True Then
    lblCD.Caption = "O"
    Velha = Velha + 1
End If
If lblCE.Caption = "X" And lblCC.Caption = "X" And lblCD.Caption = "X" Then
    LineHC.Visible = True
    Velha = 0
End If
If lblCE.Caption = "X" And lblCC.Caption = "X" And lblCD.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblSD.Caption = "X" And lblCD.Caption = "X" And lblID.Caption = "X" Then
    LineVD.Visible = True
    Velha = 0
End If
If lblSD.Caption = "X" And lblCD.Caption = "X" And lblID.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblCE.Caption = "O" And lblCC.Caption = "O" And lblCD.Caption = "O" Then
    LineHC.Visible = True
    Velha = 0
End If
If lblCE.Caption = "O" And lblCC.Caption = "O" And lblCD.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
    If lblSD.Caption = "O" And lblCD.Caption = "O" And lblID.Caption = "O" Then
    LineVD.Visible = True
    Velha = 0
End If
If lblSD.Caption = "O" And lblCD.Caption = "O" And lblID.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblCD.Caption = "X" Then
    OptionX.Enabled = False
    OptionO.Enabled = True
    OptionX.Value = False
    OptionO.Value = True
    Else
    OptionX.Enabled = True
    OptionO.Enabled = False
    OptionX.Value = True
    OptionO.Value = False
End If
If Velha = 9 Then
    lblRV.Caption = lblRV.Caption + 1
End If
If lblCD.Caption = "X" Or lblCD.Caption = "O" Then
    lblCD.Enabled = False
End If
If LineHS.Visible = True Or LineHC.Visible = True Or LineHI.Visible = True Or LineVE.Visible = True Or LineVC.Visible = True Or LineVD.Visible = True Or LineDE.Visible = True Or LineDD.Visible = True Then
    lblSE.Enabled = False
    lblSC.Enabled = False
    lblSD.Enabled = False
    lblIE.Enabled = False
    lblIC.Enabled = False
    lblID.Enabled = False
    lblCE.Enabled = False
    lblCC.Enabled = False
    lblCD.Enabled = False
End If
End Sub

Private Sub lblCE_Click()
If OptionX.Value = True Then
    lblCE.Caption = "X"
    Velha = Velha + 1
    ElseIf OptionO.Value = True Then
    lblCE.Caption = "O"
    Velha = Velha + 1
End If
If lblCE.Caption = "X" And lblCC.Caption = "X" And lblCD.Caption = "X" Then
    LineHC.Visible = True
    Velha = 0
End If
If lblSE.Caption = "X" And lblCE.Caption = "X" And lblIE.Caption = "X" Then
    LineVE.Visible = True
    Velha = 0
End If
If lblCE.Caption = "X" And lblCC.Caption = "X" And lblCD.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblSE.Caption = "X" And lblCE.Caption = "X" And lblIE.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblCE.Caption = "O" And lblCC.Caption = "O" And lblCD.Caption = "O" Then
    LineHC.Visible = True
    Velha = 0
End If
If lblSE.Caption = "O" And lblCE.Caption = "O" And lblIE.Caption = "O" Then
    LineVE.Visible = True
    Velha = 0
End If
If lblCE.Caption = "O" And lblCC.Caption = "O" And lblCD.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblSE.Caption = "O" And lblCE.Caption = "O" And lblIE.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblCE.Caption = "X" Then
    OptionX.Enabled = False
    OptionO.Enabled = True
    OptionX.Value = False
    OptionO.Value = True
    Else
    OptionX.Enabled = True
    OptionO.Enabled = False
    OptionX.Value = True
    OptionO.Value = False
End If
If Velha = 9 Then
    lblRV.Caption = lblRV.Caption + 1
End If
If lblCE.Caption = "X" Or lblCE.Caption = "O" Then
    lblCE.Enabled = False
End If
If LineHS.Visible = True Or LineHC.Visible = True Or LineHI.Visible = True Or LineVE.Visible = True Or LineVC.Visible = True Or LineVD.Visible = True Or LineDE.Visible = True Or LineDD.Visible = True Then
    lblSE.Enabled = False
    lblSC.Enabled = False
    lblSD.Enabled = False
    lblIE.Enabled = False
    lblIC.Enabled = False
    lblID.Enabled = False
    lblCE.Enabled = False
    lblCC.Enabled = False
    lblCD.Enabled = False
End If
End Sub

Private Sub lblIC_Click()
If OptionX.Value = True Then
    lblIC.Caption = "X"
    Velha = Velha + 1
    ElseIf OptionO.Value = True Then
    lblIC.Caption = "O"
    Velha = Velha + 1
End If
If lblIE.Caption = "X" And lblIC.Caption = "X" And lblID.Caption = "X" Then
    LineHI.Visible = True
    Velha = 0
End If
If lblCC.Caption = "X" And lblSC.Caption = "X" And lblIC.Caption = "X" Then
    LineVC.Visible = True
    Velha = 0
End If
If lblIE.Caption = "X" And lblIC.Caption = "X" And lblID.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblSC.Caption = "X" And lblCC.Caption = "X" And lblIC.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblIE.Caption = "O" And lblIC.Caption = "O" And lblID.Caption = "O" Then
    LineHI.Visible = True
    Velha = 0
End If
If lblCC.Caption = "O" And lblSC.Caption = "O" And lblIC.Caption = "O" Then
    LineVC.Visible = True
    Velha = 0
End If
If lblIE.Caption = "O" And lblIC.Caption = "O" And lblID.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblSC.Caption = "O" And lblCC.Caption = "O" And lblIC.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblIC.Caption = "X" Then
    OptionX.Enabled = False
    OptionO.Enabled = True
    OptionX.Value = False
    OptionO.Value = True
    Else
    OptionX.Enabled = True
    OptionO.Enabled = False
    OptionX.Value = True
    OptionO.Value = False
End If
If Velha = 9 Then
    lblRV.Caption = lblRV.Caption + 1
End If
If lblIC.Caption = "X" Or lblIC.Caption = "O" Then
    lblIC.Enabled = False
End If
If LineHS.Visible = True Or LineHC.Visible = True Or LineHI.Visible = True Or LineVE.Visible = True Or LineVC.Visible = True Or LineVD.Visible = True Or LineDE.Visible = True Or LineDD.Visible = True Then
    lblSE.Enabled = False
    lblSC.Enabled = False
    lblSD.Enabled = False
    lblIE.Enabled = False
    lblIC.Enabled = False
    lblID.Enabled = False
    lblCE.Enabled = False
    lblCC.Enabled = False
    lblCD.Enabled = False
End If
End Sub

Private Sub lblID_Click()
If OptionX.Value = True Then
    lblID.Caption = "X"
    Velha = Velha + 1
    ElseIf OptionO.Value = True Then
    lblID.Caption = "O"
    Velha = Velha + 1
End If
If lblIE.Caption = "X" And lblIC.Caption = "X" And lblID.Caption = "X" Then
    LineHI.Visible = True
    Velha = 0
End If
If lblSD.Caption = "X" And lblCD.Caption = "X" And lblID.Caption = "X" Then
    LineVD.Visible = True
    Velha = 0
End If
If lblID.Caption = "X" And lblCC.Caption = "X" And lblSE.Caption = "X" Then
    LineDE.Visible = True
    Velha = 0
End If
If lblIE.Caption = "X" And lblIC.Caption = "X" And lblID.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblID.Caption = "X" And lblCC.Caption = "X" And lblSE.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblSD.Caption = "X" And lblCD.Caption = "X" And lblID.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblIE.Caption = "O" And lblIC.Caption = "O" And lblID.Caption = "O" Then
    LineHI.Visible = True
    Velha = 0
End If
If lblSD.Caption = "O" And lblCD.Caption = "O" And lblID.Caption = "O" Then
    LineVD.Visible = True
    Velha = 0
End If
If lblID.Caption = "O" And lblCC.Caption = "O" And lblSE.Caption = "O" Then
    LineDE.Visible = True
    Velha = 0
End If
If lblIE.Caption = "O" And lblIC.Caption = "O" And lblID.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblID.Caption = "O" And lblCC.Caption = "O" And lblSE.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblSD.Caption = "O" And lblCD.Caption = "O" And lblID.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblID.Caption = "X" Then
    OptionX.Enabled = False
    OptionO.Enabled = True
    OptionX.Value = False
    OptionO.Value = True
    Else
    OptionX.Enabled = True
    OptionO.Enabled = False
    OptionX.Value = True
    OptionO.Value = False
End If
If Velha = 9 Then
    lblRV.Caption = lblRV.Caption + 1
End If
If lblID.Caption = "X" Or lblID.Caption = "O" Then
    lblID.Enabled = False
End If
If LineHS.Visible = True Or LineHC.Visible = True Or LineHI.Visible = True Or LineVE.Visible = True Or LineVC.Visible = True Or LineVD.Visible = True Or LineDE.Visible = True Or LineDD.Visible = True Then
    lblSE.Enabled = False
    lblSC.Enabled = False
    lblSD.Enabled = False
    lblIE.Enabled = False
    lblIC.Enabled = False
    lblID.Enabled = False
    lblCE.Enabled = False
    lblCC.Enabled = False
    lblCD.Enabled = False
End If
End Sub

Private Sub lblIE_Click()
If OptionX.Value = True Then
    lblIE.Caption = "X"
    Velha = Velha + 1
    ElseIf OptionO.Value = True Then
    lblIE.Caption = "O"
    Velha = Velha + 1
End If
If lblIE.Caption = "X" And lblIC.Caption = "X" And lblID.Caption = "X" Then
    LineHI.Visible = True
    Velha = 0
End If
If lblSE.Caption = "X" And lblCE.Caption = "X" And lblIE.Caption = "X" Then
    LineVE.Visible = True
    Velha = 0
End If
If lblIE.Caption = "X" And lblCC.Caption = "X" And lblSD.Caption = "X" Then
    LineDD.Visible = True
    Velha = 0
End If
If lblIE.Caption = "X" And lblIC.Caption = "X" And lblID.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblSE.Caption = "X" And lblCE.Caption = "X" And lblIE.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblIE.Caption = "X" And lblCC.Caption = "X" And lblSD.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblIE.Caption = "O" And lblIC.Caption = "O" And lblID.Caption = "O" Then
    LineHI.Visible = True
    Velha = 0
End If
If lblSE.Caption = "O" And lblCE.Caption = "O" And lblIE.Caption = "O" Then
    LineVE.Visible = True
    Velha = 0
End If
If lblIE.Caption = "O" And lblCC.Caption = "O" And lblSD.Caption = "O" Then
    LineDD.Visible = True
    Velha = 0
End If
If lblIE.Caption = "O" And lblIC.Caption = "O" And lblID.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblIE.Caption = "O" And lblCC.Caption = "O" And lblSD.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblSE.Caption = "O" And lblCE.Caption = "O" And lblIE.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
    Velha = 0
End If
If lblIE.Caption = "X" Then
    OptionX.Enabled = False
    OptionO.Enabled = True
    OptionX.Value = False
    OptionO.Value = True
    Else
    OptionX.Enabled = True
    OptionO.Enabled = False
    OptionX.Value = True
    OptionO.Value = False
End If
If Velha = 9 Then
    lblRV.Caption = lblRV.Caption + 1
End If
If lblIE.Caption = "X" Or lblIE.Caption = "O" Then
    lblIE.Enabled = False
End If
If LineHS.Visible = True Or LineHC.Visible = True Or LineHI.Visible = True Or LineVE.Visible = True Or LineVC.Visible = True Or LineVD.Visible = True Or LineDE.Visible = True Or LineDD.Visible = True Then
    lblSE.Enabled = False
    lblSC.Enabled = False
    lblSD.Enabled = False
    lblIE.Enabled = False
    lblIC.Enabled = False
    lblID.Enabled = False
    lblCE.Enabled = False
    lblCC.Enabled = False
    lblCD.Enabled = False
End If
End Sub
Private Sub lblSC_Click()
If OptionX.Value = True Then
    lblSC.Caption = "X"
    Velha = Velha + 1
    ElseIf OptionO.Value = True Then
    lblSC.Caption = "O"
    Velha = Velha + 1
End If
If lblSE.Caption = "X" And lblSC.Caption = "X" And lblSD.Caption = "X" Then
    LineHS.Visible = True
    Velha = 0
End If
If lblCC.Caption = "X" And lblSC.Caption = "X" And lblIC.Caption = "X" Then
    LineVC.Visible = True
    Velha = 0
End If
If lblSE.Caption = "X" And lblSC.Caption = "X" And lblSD.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblSC.Caption = "X" And lblCC.Caption = "X" And lblIC.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblSE.Caption = "O" And lblSC.Caption = "O" And lblSD.Caption = "O" Then
    LineHS.Visible = True
    Velha = 0
End If
If lblCC.Caption = "O" And lblSC.Caption = "O" And lblIC.Caption = "O" Then
    LineVC.Visible = True
    Velha = 0
End If
If lblSE.Caption = "O" And lblSC.Caption = "O" And lblSD.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblSC.Caption = "O" And lblCC.Caption = "O" And lblIC.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblSC.Caption = "X" Then
    OptionX.Enabled = False
    OptionO.Enabled = True
    OptionX.Value = False
    OptionO.Value = True
    Else
    OptionX.Enabled = True
    OptionO.Enabled = False
    OptionX.Value = True
    OptionO.Value = False
End If
If Velha = 9 Then
    lblRV.Caption = lblRV.Caption + 1
End If
If lblSC.Caption = "X" Or lblSC.Caption = "O" Then
    lblSC.Enabled = False
End If
If LineHS.Visible = True Or LineHC.Visible = True Or LineHI.Visible = True Or LineVE.Visible = True Or LineVC.Visible = True Or LineVD.Visible = True Or LineDE.Visible = True Or LineDD.Visible = True Then
    lblSE.Enabled = False
    lblSC.Enabled = False
    lblSD.Enabled = False
    lblIE.Enabled = False
    lblIC.Enabled = False
    lblID.Enabled = False
    lblCE.Enabled = False
    lblCC.Enabled = False
    lblCD.Enabled = False
End If
End Sub

Private Sub lblSD_Click()
If OptionX.Value = True Then
    lblSD.Caption = "X"
    Velha = Velha + 1
    ElseIf OptionO.Value = True Then
    lblSD.Caption = "O"
    Velha = Velha + 1
End If
If lblSE.Caption = "X" And lblSC.Caption = "X" And lblSD.Caption = "X" Then
    LineHS.Visible = True
    Velha = 0
End If
If lblIE.Caption = "X" And lblCC.Caption = "X" And lblSD.Caption = "X" Then
    LineDD.Visible = True
    Velha = 0
End If
If lblSD.Caption = "X" And lblCD.Caption = "X" And lblID.Caption = "X" Then
    LineVD.Visible = True
    Velha = 0
End If
If lblSE.Caption = "X" And lblSC.Caption = "X" And lblSD.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblIE.Caption = "X" And lblCC.Caption = "X" And lblSD.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblSD.Caption = "X" And lblCD.Caption = "X" And lblID.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblSE.Caption = "O" And lblSC.Caption = "O" And lblSD.Caption = "O" Then
    LineHS.Visible = True
    Velha = 0
End If
If lblIE.Caption = "O" And lblCC.Caption = "O" And lblSD.Caption = "O" Then
    LineDD.Visible = True
    Velha = 0
End If
If lblSD.Caption = "O" And lblCD.Caption = "O" And lblID.Caption = "O" Then
    LineVD.Visible = True
    Velha = 0
End If
If lblSE.Caption = "O" And lblSC.Caption = "O" And lblSD.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblIE.Caption = "O" And lblCC.Caption = "O" And lblSD.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblSD.Caption = "O" And lblCD.Caption = "O" And lblID.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblSD.Caption = "X" Then
    OptionX.Enabled = False
    OptionO.Enabled = True
    OptionX.Value = False
    OptionO.Value = True
    Else
    OptionX.Enabled = True
    OptionO.Enabled = False
    OptionX.Value = True
    OptionO.Value = False
End If
If Velha = 9 Then
    lblRV.Caption = lblRV.Caption + 1
End If
If lblSD.Caption = "X" Or lblSD.Caption = "O" Then
    lblSD.Enabled = False
End If
If LineHS.Visible = True Or LineHC.Visible = True Or LineHI.Visible = True Or LineVE.Visible = True Or LineVC.Visible = True Or LineVD.Visible = True Or LineDE.Visible = True Or LineDD.Visible = True Then
    lblSE.Enabled = False
    lblSC.Enabled = False
    lblSD.Enabled = False
    lblIE.Enabled = False
    lblIC.Enabled = False
    lblID.Enabled = False
    lblCE.Enabled = False
    lblCC.Enabled = False
    lblCD.Enabled = False
End If
End Sub

Private Sub lblSE_Click()
If OptionX.Value = True Then
    lblSE.Caption = "X"
    Velha = Velha + 1
    ElseIf OptionO.Value = True Then
        lblSE.Caption = "O"
        Velha = Velha + 1
End If
If lblSE.Caption = "X" And lblSC.Caption = "X" And lblSD.Caption = "X" Then
    LineHS.Visible = True
    Velha = 0
End If
If lblSE.Caption = "X" And lblCE.Caption = "X" And lblIE.Caption = "X" Then
    LineVE.Visible = True
    Velha = 0
End If
If lblID.Caption = "X" And lblCC.Caption = "X" And lblSE.Caption = "X" Then
    LineDE.Visible = True
    Velha = 0
End If
If lblSE.Caption = "X" And lblSC.Caption = "X" And lblSD.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblSE.Caption = "X" And lblCE.Caption = "X" And lblIE.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblID.Caption = "X" And lblCC.Caption = "X" And lblSE.Caption = "X" Then
    lblRX.Caption = lblRX.Caption + 1
End If
If lblSE.Caption = "O" And lblSC.Caption = "O" And lblSD.Caption = "O" Then
    LineHS.Visible = True
    Velha = 0
End If
If lblSE.Caption = "O" And lblCE.Caption = "O" And lblIE.Caption = "O" Then
    LineVE.Visible = True
    Velha = 0
End If
If lblID.Caption = "O" And lblCC.Caption = "O" And lblSE.Caption = "O" Then
    LineDE.Visible = True
    Velha = 0
End If
If lblSE.Caption = "O" And lblSC.Caption = "O" And lblSD.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblSE.Caption = "O" And lblCE.Caption = "O" And lblIE.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblID.Caption = "O" And lblCC.Caption = "O" And lblSE.Caption = "O" Then
    lblRO.Caption = lblRO.Caption + 1
End If
If lblSE.Caption = "X" Then
    OptionX.Enabled = False
    OptionO.Enabled = True
    OptionX.Value = False
    OptionO.Value = True
    Else
    OptionX.Enabled = True
    OptionO.Enabled = False
    OptionX.Value = True
    OptionO.Value = False
End If
If Velha = 9 Then
    lblRV.Caption = lblRV.Caption + 1
End If
If lblSE.Caption = "X" Or lblSE.Caption = "O" Then
    lblSE.Enabled = False
End If
If LineHS.Visible = True Or LineHC.Visible = True Or LineHI.Visible = True Or LineVE.Visible = True Or LineVC.Visible = True Or LineVD.Visible = True Or LineDE.Visible = True Or LineDD.Visible = True Then
    lblSE.Enabled = False
    lblSC.Enabled = False
    lblSD.Enabled = False
    lblIE.Enabled = False
    lblIC.Enabled = False
    lblID.Enabled = False
    lblCE.Enabled = False
    lblCC.Enabled = False
    lblCD.Enabled = False
End If
End Sub

Private Sub OptionO_Click()
If OptionO.Enabled = True Then
    OptionX.Enabled = False
End If
End Sub

Private Sub OptionX_Click()
If OptionX.Enabled = True Then
    OptionO.Enabled = False
End If
End Sub
