VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jogo da Velha"
   ClientHeight    =   5730
   ClientLeft      =   8640
   ClientTop       =   3495
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNew 
      Caption         =   "Novo Jogo"
      Height          =   615
      Left            =   840
      TabIndex        =   21
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdZero 
      Caption         =   "Zerar placar"
      Height          =   615
      Left            =   840
      TabIndex        =   12
      Top             =   4920
      Width           =   1335
   End
   Begin VB.OptionButton optO 
      BackColor       =   &H0080FF80&
      Caption         =   "O"
      Height          =   255
      Left            =   5280
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4680
      Width           =   495
   End
   Begin VB.OptionButton optX 
      BackColor       =   &H0080FF80&
      Caption         =   "X"
      Height          =   255
      Left            =   3960
      TabIndex        =   0
      Top             =   4680
      Value           =   -1  'True
      Width           =   495
   End
   Begin VB.Line Line11 
      BorderWidth     =   3
      X1              =   5160
      X2              =   6840
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line10 
      BorderWidth     =   3
      X1              =   5160
      X2              =   6840
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line9 
      BorderWidth     =   3
      X1              =   6840
      X2              =   6840
      Y1              =   1200
      Y2              =   2160
   End
   Begin VB.Line Line8 
      BorderWidth     =   3
      X1              =   5160
      X2              =   5160
      Y1              =   1200
      Y2              =   2160
   End
   Begin VB.Label lblAtual 
      BackColor       =   &H0080FF80&
      Caption         =   "Jogador atual: X"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   20
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Line LineDDE 
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   3240
      X2              =   720
      Y1              =   2640
      Y2              =   120
   End
   Begin VB.Line LineDED 
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   720
      X2              =   3240
      Y1              =   2640
      Y2              =   120
   End
   Begin VB.Line LineVE 
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   1200
      X2              =   1200
      Y1              =   2640
      Y2              =   120
   End
   Begin VB.Line LineVC 
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   1920
      X2              =   1920
      Y1              =   2640
      Y2              =   120
   End
   Begin VB.Line LineVD 
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   2760
      X2              =   2760
      Y1              =   2640
      Y2              =   120
   End
   Begin VB.Line LineHS 
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   3240
      X2              =   720
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line LineHC 
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   3240
      X2              =   720
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line LineHI 
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   3240
      X2              =   720
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblPlacarVelha 
      BackColor       =   &H0080FF80&
      Caption         =   "    0"
      Height          =   255
      Left            =   6240
      TabIndex        =   19
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblPlacarO 
      BackColor       =   &H0080FF80&
      Caption         =   "  0"
      Height          =   255
      Left            =   5760
      TabIndex        =   18
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label lblPlacarX 
      BackColor       =   &H0080FF80&
      Caption         =   "  0"
      Height          =   255
      Left            =   5280
      TabIndex        =   17
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label lblVelha 
      BackColor       =   &H0080FF80&
      Caption         =   "Velha"
      Height          =   255
      Left            =   6240
      TabIndex        =   16
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblO 
      BackColor       =   &H0080FF80&
      Caption         =   " O"
      Height          =   255
      Left            =   5760
      TabIndex        =   15
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lblX 
      BackColor       =   &H0080FF80&
      Caption         =   "  X"
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      Top             =   1320
      Width           =   255
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   6120
      X2              =   6120
      Y1              =   1200
      Y2              =   2160
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   5640
      X2              =   5640
      Y1              =   1200
      Y2              =   2160
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   5160
      X2              =   6840
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblPlacar 
      BackColor       =   &H0080FF80&
      Caption         =   "    Placar"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   13
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblSE 
      BackColor       =   &H0080FF80&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   840
      TabIndex        =   11
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblCE 
      BackColor       =   &H0080FF80&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   10
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblIE 
      BackColor       =   &H0080FF80&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   9
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblIC 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1680
      TabIndex        =   8
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblC 
      BackColor       =   &H0080FF80&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   7
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblSC 
      BackColor       =   &H0080FF80&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   6
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblID 
      BackColor       =   &H0080FF80&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblCD 
      BackColor       =   &H0080FF80&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblSD 
      BackColor       =   &H0080FF80&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      X1              =   840
      X2              =   3120
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      X1              =   840
      X2              =   3120
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderWidth     =   10
      X1              =   2400
      X2              =   2400
      Y1              =   240
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      X1              =   1560
      X2              =   1560
      Y1              =   240
      Y2              =   2520
   End
   Begin VB.Label lblLance 
      BackColor       =   &H0080FF80&
      Caption         =   "Primeiro lance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   2
      Top             =   3720
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CHEIOCD, CHEIOSE As Boolean
Dim CHEIOSC, CHEIOID As Boolean
Dim CHEIOCE, CHEIOSD As Boolean
Dim VELHAPT, PRIMJOG As Integer
Dim CHEIOIE, ENDGAME As Boolean
Dim CHEIOIC, CHEIOC As Boolean
Dim JOOJX, JOOJO As Boolean

Private Sub cmdNew_Click()
PRIMJOG = 0
VELHAPT = 0
cmdNew.Enabled = False
CHEIOSE = False
CHEIOSC = False
CHEIOSD = False
CHEIOCD = False
CHEIOC = False
ENDGAME = False
optX.Enabled = True
optO.Enabled = True
optX.Value = False
optO.Value = False
CHEIOCE = False
CHEIOIE = False
CHEIOIC = False
CHEIOID = False
LineVC.Visible = False
LineVE.Visible = False
LineVD.Visible = False
LineHS.Visible = False
LineHI.Visible = False
LineHC.Visible = False
LineDDE.Visible = False
LineDED.Visible = False
lblSE = ""
lblSC = ""
lblSD = ""
lblCD = ""
lblC = ""
lblCE = ""
lblIE = ""
lblIC = ""
lblID = ""
If lblAtual = "Jogador atual: X" Then
    optX.Value = True
Else
    optO.Value = True
End If
If Int(lblPlacarO.Caption) > 0 Or Int(lblPlacarX.Caption) > 0 Or Int(lblPlacarVelha.Caption) > 0 Then
    cmdZero.Enabled = True
End If
End Sub

Private Sub cmdZero_Click()
lblPlacarX.Caption = "  0"
lblPlacarO.Caption = "  0"
lblPlacarVelha.Caption = "  0"
cmdZero.Enabled = False
End Sub

Private Sub Form_Load()
cmdZero.Enabled = False
cmdNew.Enabled = False
JOOJX = True
PRIMJOG = 0
End Sub

Private Sub lblC_Click()
If ENDGAME = False Then
    cmdNew.Enabled = True
    If CHEIOC = False Then
        VELHAPT = VELHAPT + 1
        If JOOJX = True Then
            lblC.ForeColor = &HFF&
            lblC.Caption = "  X"
            JOOJX = False
            JOOJO = True
            CHEIOC = True
        Else
            If JOOJO = True Then
                lblC.ForeColor = &HFF0000
                lblC.Caption = "  O"
                JOOJO = False
                JOOJX = True
                CHEIOC = True
            End If
        End If
    Else
            If CHEIOC = True Then
                Call MsgBox("Selecione outro espaço")
            End If
    End If
End If
PRIMJOG = 1
If CHEIOSC = True And CHEIOC = True And CHEIOIC = True Then
    If lblSC = "  X" And lblC = "  X" And lblIC = "  X" Then
        LineVC.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblSC = "  O" And lblC = "  O" And lblIC = "  O" Then
        LineVC.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If CHEIOSE = True And CHEIOC = True And CHEIOID = True And VELHAPT > 0 Then
    If lblSE = "  X" And lblC = "  X" And lblID = "  X" Then
        LineDDE.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblSE = "  O" And lblC = "  O" And lblID = "  O" Then
        LineDDE.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If CHEIOCE = True And CHEIOC = True And CHEIOCD = True And VELHAPT > 0 Then
    If lblCE = "  X" And lblC = "  X" And lblCD = "  X" Then
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        LineHC.Visible = True
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblCE = "  O" And lblC = "  O" And lblCD = "  O" Then
        LineHC.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If CHEIOSD = True And CHEIOC = True And CHEIOIE = True And VELHAPT > 0 Then
    If lblSD = "  X" And lblC = "  X" And lblIE = "  X" Then
        LineDED.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblSD = "  O" And lblC = "  O" And lblIE = "  O" Then
        LineDED.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If VELHAPT = 9 And ENDGAME = False Then
    lblPlacarVelha.Caption = lblPlacarVelha.Caption + 1
    ENDGAME = True
    cmdZero.Enabled = True
    MsgBox ("Empate!!! A velha ganhou desta vez")
End If
If JOOJX = True Then
    lblAtual.Caption = "Jogador atual: X"
End If
If JOOJX = False Then
    lblAtual.Caption = "Jogador atual: O"
End If
End Sub

Private Sub lblCD_Click()
If ENDGAME = False Then
    cmdNew.Enabled = True
    If CHEIOCD = False Then
        VELHAPT = VELHAPT + 1
        If JOOJX = True Then
            lblCD.ForeColor = &HFF&
            lblCD.Caption = "  X"
            JOOJX = False
            JOOJO = True
            CHEIOCD = True
        Else
            If JOOJO = True Then
                lblCD.ForeColor = &HFF0000
                lblCD.Caption = "  O"
                JOOJO = False
                JOOJX = True
                CHEIOCD = True
            End If
        End If
        Else
            If CHEIOCD = True Then
                Call MsgBox("Selecione outro espaço")
            End If
    End If
End If
PRIMJOG = 1
If CHEIOCE = True And CHEIOC = True And CHEIOCD = True Then
    If lblCE = "  X" And lblC = "  X" And lblCD = "  X" Then
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        LineHC.Visible = True
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblCE = "  O" And lblC = "  O" And lblCD = "  O" Then
        LineHC.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If CHEIOSD = True And CHEIOCD = True And CHEIOID = True And VELHAPT > 0 Then
    If lblSD = "  X" And lblCD = "  X" And lblID = "  X" Then
        LineVD.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblSD = "  O" And lblCD = "  O" And lblID = "  O" Then
        LineVD.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If VELHAPT = 9 And ENDGAME = False Then
    lblPlacarVelha.Caption = lblPlacarVelha.Caption + 1
    ENDGAME = True
    cmdZero.Enabled = True
    MsgBox ("Empate!!! A velha ganhou desta vez")
End If
If JOOJX = True Then
    lblAtual.Caption = "Jogador atual: X"
End If
If JOOJX = False Then
    lblAtual.Caption = "Jogador atual: O"
End If
End Sub

Private Sub lblCE_Click()
If ENDGAME = False Then
    cmdNew.Enabled = True
    If CHEIOCE = False Then
        VELHAPT = VELHAPT + 1
        If JOOJX = True Then
            lblCE.ForeColor = &HFF&
            lblCE.Caption = "  X"
            JOOJX = False
            JOOJO = True
            CHEIOCE = True
        Else
            If JOOJO = True Then
                lblCE.ForeColor = &HFF0000
                lblCE.Caption = "  O"
                JOOJO = False
                JOOJX = True
                CHEIOCE = True
            End If
        End If
        Else
            If CHEIOCE = True Then
                Call MsgBox("Selecione outro espaço")
            End If
    End If
End If
PRIMJOG = 1
If CHEIOCE = True And CHEIOC = True And CHEIOCD = True Then
    If lblCE = "  X" And lblC = "  X" And lblCD = "  X" Then
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        LineHC.Visible = True
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblCE = "  O" And lblC = "  O" And lblCD = "  O" Then
        LineHC.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If CHEIOSE = True And CHEIOCE = True And CHEIOIE = True And VELHAPT > 0 Then
    If lblSE = "  X" And lblCE = "  X" And lblIE = "  X" Then
        LineVE.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblSE = "  O" And lblCE = "  O" And lblIE = "  O" Then
        LineVE.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If VELHAPT = 9 And ENDGAME = False Then
    lblPlacarVelha.Caption = lblPlacarVelha.Caption + 1
    ENDGAME = True
    cmdZero.Enabled = True
    MsgBox ("Empate!!! A velha ganhou desta vez")
End If
If JOOJX = True Then
    lblAtual.Caption = "Jogador atual: X"
End If
If JOOJX = False Then
    lblAtual.Caption = "Jogador atual: O"
End If
End Sub

Private Sub lblIC_Click()
If ENDGAME = False Then
    cmdNew.Enabled = True
    If CHEIOIC = False Then
        VELHAPT = VELHAPT + 1
        If JOOJX = True Then
            lblIC.ForeColor = &HFF&
            lblIC.Caption = "  X"
            JOOJX = False
            JOOJO = True
            CHEIOIC = True
        Else
            If JOOJO = True Then
                lblIC.ForeColor = &HFF0000
                lblIC.Caption = "  O"
                JOOJO = False
                JOOJX = True
                CHEIOIC = True
            End If
        End If
        Else
            If CHEIOIC = True Then
                Call MsgBox("Selecione outro espaço")
            End If
    End If
End If
PRIMJOG = 1
If CHEIOIC = True And CHEIOID = True And CHEIOIE = True Then
    If lblIE = "  X" And lblIC = "  X" And lblID = "  X" Then
        LineHI.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblIE = "  O" And lblIC = "  O" And lblID = "  O" Then
        LineHI.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If CHEIOSC = True And CHEIOC = True And CHEIOIC = True And VELHAPT > 0 Then
    If lblSC = "  X" And lblC = "  X" And lblIC = "  X" Then
        LineVC.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblSC = "  O" And lblC = "  O" And lblIC = "  O" Then
        LineVC.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If VELHAPT = 9 And ENDGAME = False Then
    lblPlacarVelha.Caption = lblPlacarVelha.Caption + 1
    ENDGAME = True
    cmdZero.Enabled = True
    MsgBox ("Empate!!! A velha ganhou desta vez")
End If
If JOOJX = True Then
    lblAtual.Caption = "Jogador atual: X"
End If
If JOOJX = False Then
    lblAtual.Caption = "Jogador atual: O"
End If
End Sub

Private Sub lblID_Click()
If ENDGAME = False Then
    cmdNew.Enabled = True
    If CHEIOID = False Then
        VELHAPT = VELHAPT + 1
        If JOOJX = True Then
            lblID.ForeColor = &HFF&
            lblID.Caption = "  X"
            JOOJX = False
            JOOJO = True
            CHEIOID = True
        Else
            If JOOJO = True Then
                lblID.ForeColor = &HFF0000
                lblID.Caption = "  O"
                JOOJO = False
                JOOJX = True
                CHEIOID = True
            End If
        End If
        Else
            If CHEIOID = True Then
                Call MsgBox("Selecione outro espaço")
            End If
    End If
End If
PRIMJOG = 1
If CHEIOIC = True And CHEIOID = True And CHEIOIE = True Then
    If lblIE = "  X" And lblIC = "  X" And lblID = "  X" Then
        LineHI.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblIE = "  O" And lblIC = "  O" And lblID = "  O" Then
        LineHI.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If CHEIOSD = True And CHEIOCD = True And CHEIOID = True And VELHAPT > 0 Then
    If lblSD = "  X" And lblCD = "  X" And lblID = "  X" Then
        LineVD.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblSD = "  O" And lblCD = "  O" And lblID = "  O" Then
        LineVD.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If CHEIOSE = True And CHEIOC = True And CHEIOID = True And VELHAPT > 0 Then
    If lblSE = "  X" And lblC = "  X" And lblID = "  X" Then
        LineDDE.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblSE = "  O" And lblC = "  O" And lblID = "  O" Then
        LineDDE.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If VELHAPT = 9 And ENDGAME = False Then
    lblPlacarVelha.Caption = lblPlacarVelha.Caption + 1
    ENDGAME = True
    cmdZero.Enabled = True
    MsgBox ("Empate!!! A velha ganhou desta vez")
End If
If JOOJX = True Then
    lblAtual.Caption = "Jogador atual: X"
End If
If JOOJX = False Then
    lblAtual.Caption = "Jogador atual: O"
End If
End Sub

Private Sub lblIE_Click()
If ENDGAME = False Then
    cmdNew.Enabled = True
    If CHEIOIE = False Then
        VELHAPT = VELHAPT + 1
        If JOOJX = True Then
            lblIE.ForeColor = &HFF&
            lblIE.Caption = "  X"
            JOOJX = False
            JOOJO = True
            CHEIOIE = True
        Else
            If JOOJO = True Then
                lblIE.ForeColor = &HFF0000
                lblIE.Caption = "  O"
                JOOJO = False
                JOOJX = True
                CHEIOIE = True
            End If
        End If
        Else
            If CHEIOIE = True Then
                Call MsgBox("Selecione outro espaço")
            End If
    End If
End If
PRIMJOG = 1
If CHEIOIC = True And CHEIOID = True And CHEIOIE = True Then
    If lblIE = "  X" And lblIC = "  X" And lblID = "  X" Then
        LineHI.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblIE = "  O" And lblIC = "  O" And lblID = "  O" Then
        LineHI.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If CHEIOSE = True And CHEIOCE = True And CHEIOIE = True And VELHAPT > 0 Then
    If lblSE = "  X" And lblCE = "  X" And lblIE = "  X" Then
        LineVE.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblSE = "  O" And lblCE = "  O" And lblIE = "  O" Then
        LineVE.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If CHEIOSD = True And CHEIOC = True And CHEIOIE = True And VELHAPT > 0 Then
    If lblSD = "  X" And lblC = "  X" And lblIE = "  X" Then
        LineDED.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblSD = "  O" And lblC = "  O" And lblIE = "  O" Then
        LineDED.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If VELHAPT = 9 And ENDGAME = False Then
    lblPlacarVelha.Caption = lblPlacarVelha.Caption + 1
    ENDGAME = True
    cmdZero.Enabled = True
    MsgBox ("Empate!!! A velha ganhou desta vez")
End If
If JOOJX = True Then
    lblAtual.Caption = "Jogador atual: X"
End If
If JOOJX = False Then
    lblAtual.Caption = "Jogador atual: O"
End If
End Sub




Private Sub lblSC_Click()
If ENDGAME = False Then
    cmdNew.Enabled = True
    If CHEIOSC = False Then
        VELHAPT = VELHAPT + 1
        If JOOJX = True Then
            lblSC.ForeColor = &HFF&
            lblSC.Caption = "  X"
            JOOJX = False
            JOOJO = True
            CHEIOSC = True
        Else
            If JOOJO = True Then
                lblSC.ForeColor = &HFF0000
                lblSC.Caption = "  O"
                JOOJO = False
                JOOJX = True
                CHEIOSC = True
            End If
        End If
        Else
            If CHEIOSC = True Then
                Call MsgBox("Selecione outro espaço")
            End If
    End If
End If
PRIMJOG = 1
If CHEIOSE And CHEIOSC And CHEIOSD = True Then
    If lblSE = "  X" And lblSC = "  X" And lblSD = "  X" Then
        LineHS.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblSE = "  O" And lblSC = "  O" And lblSD = "  O" Then
        LineHS.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If CHEIOSC = True And CHEIOC = True And CHEIOIC = True And VELHAPT > 0 Then
    If lblSC = "  X" And lblC = "  X" And lblIC = "  X" Then
        LineVC.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblSC = "  O" And lblC = "  O" And lblIC = "  O" Then
        LineVC.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If VELHAPT = 9 And ENDGAME = False Then
    lblPlacarVelha.Caption = lblPlacarVelha.Caption + 1
    ENDGAME = True
    cmdZero.Enabled = True
    MsgBox ("Empate!!! A velha ganhou desta vez")
End If
If JOOJX = True Then
    lblAtual.Caption = "Jogador atual: X"
End If
If JOOJX = False Then
    lblAtual.Caption = "Jogador atual: O"
End If
End Sub

Private Sub lblSD_Click()
If ENDGAME = False Then
    cmdNew.Enabled = True
    If CHEIOSD = False Then
        VELHAPT = VELHAPT + 1
        If JOOJX = True Then
            lblSD.ForeColor = &HFF&
            lblSD.Caption = "  X"
            JOOJX = False
            JOOJO = True
            CHEIOSD = True
        Else
            If JOOJO = True Then
                lblSD.ForeColor = &HFF0000
                lblSD.Caption = "  O"
                JOOJO = False
                JOOJX = True
                CHEIOSD = True
            End If
        End If
        Else
            If CHEIOSD = True Then
                Call MsgBox("Selecione outro espaço")
            End If
    End If
End If
PRIMJOG = 1
If CHEIOSE And CHEIOSC And CHEIOSD = True Then
    If lblSE = "  X" And lblSC = "  X" And lblSD = "  X" Then
        LineHS.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblSE = "  O" And lblSC = "  O" And lblSD = "  O" Then
        LineHS.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If CHEIOSD = True And CHEIOCD = True And CHEIOID = True And VELHAPT > 0 Then
    If lblSD = "  X" And lblCD = "  X" And lblID = "  X" Then
        LineVD.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblSD = "  O" And lblCD = "  O" And lblID = "  O" Then
        LineVD.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If CHEIOSD = True And CHEIOC = True And CHEIOIE = True And VELHAPT > 0 Then
    If lblSD = "  X" And lblC = "  X" And lblIE = "  X" Then
        LineDED.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblSD = "  O" And lblC = "  O" And lblIE = "  O" Then
        LineDED.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If VELHAPT = 9 And ENDGAME = False Then
    lblPlacarVelha.Caption = lblPlacarVelha.Caption + 1
    ENDGAME = True
    cmdZero.Enabled = True
    MsgBox ("Empate!!! A velha ganhou desta vez")
End If
If JOOJX = True Then
    lblAtual.Caption = "Jogador atual: X"
End If
If JOOJX = False Then
    lblAtual.Caption = "Jogador atual: O"
End If
End Sub
Private Sub lblSE_Click()
If ENDGAME = False Then
    cmdNew.Enabled = True
    If CHEIOSE = False Then
        VELHAPT = VELHAPT + 1
        If JOOJX = True Then
            lblSE.ForeColor = &HFF&
            lblSE.Caption = "  X"
            JOOJX = False
            JOOJO = True
            CHEIOSE = True
        Else
            If JOOJO = True Then
                lblSE.ForeColor = &HFF0000
                lblSE.Caption = "  O"
                JOOJO = False
                JOOJX = True
                CHEIOSE = True
            End If
        End If
        Else
            If CHEIOSE = True Then
                Call MsgBox("Selecione outro espaço")
            End If
    End If
End If
PRIMJOG = 1
If CHEIOSE And CHEIOSC And CHEIOSD = True Then
    If lblSE = "  X" And lblSC = "  X" And lblSD = "  X" Then
        LineHS.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblSE = "  O" And lblSC = "  O" And lblSD = "  O" Then
        LineHS.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If CHEIOSE = True And CHEIOCE = True And CHEIOIE = True And VELHAPT > 0 Then
    If lblSE = "  X" And lblCE = "  X" And lblIE = "  X" Then
        LineVE.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblSE = "  O" And lblCE = "  O" And lblIE = "  O" Then
        LineVE.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If CHEIOSE = True And CHEIOC = True And CHEIOID = True And VELHAPT > 0 Then
    If lblSE = "  X" And lblC = "  X" And lblID = "  X" Then
        LineDDE.Visible = True
        lblPlacarX.Caption = lblPlacarX.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("X Ganhou!!!")
    End If
    If lblSE = "  O" And lblC = "  O" And lblID = "  O" Then
        LineDDE.Visible = True
        lblPlacarO.Caption = lblPlacarO.Caption + 1
        ENDGAME = True
        cmdZero.Enabled = True
        VELHAPT = 0
        MsgBox ("O Ganhou!!!")
    End If
End If
If VELHAPT = 9 And ENDGAME = False Then
    lblPlacarVelha.Caption = lblPlacarVelha.Caption + 1
    ENDGAME = True
    cmdZero.Enabled = True
    MsgBox ("Empate!!! A velha ganhou desta vez")
End If
If JOOJX = True Then
    lblAtual.Caption = "Jogador atual: X"
End If
If JOOJX = False Then
    lblAtual.Caption = "Jogador atual: O"
End If
End Sub
Private Sub optO_Click()
If PRIMJOG = 1 Then
    Call MsgBox("Apenas é possível mudar o primeiro jogador antes da primeira jogada")
    PRIMJOG = 2
    optX.Enabled = True
    optO.Enabled = False
    optX.Value = True
    optO.Value = False
End If
If PRIMJOG = 0 Then
    lblAtual.Caption = "Jogador atual: O"
    JOOJX = False
    JOOJO = True
End If
End Sub

Private Sub optX_Click()
If PRIMJOG = 1 Then
    Call MsgBox("Apenas é possível mudar o primeiro jogador antes da primeira jogada")
    PRIMJOG = 2
    optX.Enabled = False
    optO.Enabled = True
    optX.Value = False
    optO.Value = True
End If
If PRIMJOG = 0 Then
    lblAtual.Caption = "Jogador atual: X"
    JOOJX = True
    JOOJO = False
End If
End Sub
