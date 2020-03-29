VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Projeto Calculadora"
   ClientHeight    =   2955
   ClientLeft      =   8295
   ClientTop       =   3975
   ClientWidth     =   2475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   2475
   Begin VB.CommandButton cmdPi 
      Caption         =   "Pi"
      Height          =   375
      Left            =   840
      TabIndex        =   20
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton cmdR 
      Caption         =   "Raiz"
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton cmdP 
      Caption         =   "^"
      Height          =   375
      Left            =   1320
      TabIndex        =   17
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Limpar"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdD 
      Caption         =   "/"
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdMp 
      Caption         =   "*"
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdM 
      Caption         =   "-"
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton cmdS 
      Caption         =   "+"
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton cmdE 
      Caption         =   "="
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "0"
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdV 
      Caption         =   ","
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmd9 
      Caption         =   "9"
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmd8 
      Caption         =   "8"
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "7"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmd6 
      Caption         =   "6"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "5"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "4"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "3"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "2"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   495
   End
   Begin VB.Line Line3 
      X1              =   2280
      X2              =   2640
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line linE 
      X1              =   2400
      X2              =   0
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ASMR1 As Currency
Dim ASMR2 As Currency
Dim ASMR3 As Currency
Dim SUBTON As Boolean
Dim SOMAON As Boolean
Dim VIRG As Boolean
Dim SOMA As Boolean
Dim DIVI As Boolean
Dim SUBT As Boolean
Dim MULT As Boolean
Dim POTE As Boolean
Dim RAIZREAL As Boolean
Dim RAIZ As Boolean
Private Sub cmd0_Click()
lbl1.Caption = lbl1.Caption & 0
End Sub

Private Sub cmd1_Click()
lbl1.Caption = lbl1.Caption & 1
End Sub

Private Sub cmd2_Click()
lbl1.Caption = lbl1.Caption & 2
End Sub

Private Sub cmd3_Click()
lbl1.Caption = lbl1.Caption & 3
End Sub

Private Sub cmd4_Click()
lbl1.Caption = lbl1.Caption & 4
End Sub

Private Sub cmd5_Click()
lbl1.Caption = lbl1.Caption & 5
End Sub

Private Sub cmd6_Click()
lbl1.Caption = lbl1.Caption & 6
End Sub

Private Sub cmd7_Click()
lbl1.Caption = lbl1.Caption & 7
End Sub

Private Sub cmd8_Click()
lbl1.Caption = lbl1.Caption & 8
End Sub

Private Sub cmd9_Click()
lbl1.Caption = lbl1.Caption & 9
End Sub

Private Sub cmdClear_Click()
lbl1.Caption = Empty
ASMR1 = Empty
ASMR2 = Empty
ASMR3 = Empty
VIRG = False
SOMA = False
SUBT = False
DIVI = False
POTE = False
RAIZ = False
SUBTON = False
SOMAON = False
MULT = False
End Sub

Private Sub cmdD_Click()
If lbl1.Caption = Empty Then
    Call MsgBox("Insira um número")
Else
    VIRG = False
    ASMR1 = CCur(lbl1.Caption)
    lbl1.Caption = Empty
    SOMA = False
    SUBT = False
    DIVI = True
    SUBTON = False
    SOMAON = False
    POTE = False
    RAIZ = False
    MULT = False

End If
End Sub

Private Sub cmdE_Click()
    If lbl1.Caption = Empty Then
        Call MsgBox("Insira um número")
    Else
        ASMR2 = lbl1.Caption
        If IsNumeric(ASMR2) = True Then
        
            If SOMA = True Then
                VIRG = False
                lbl1.Caption = ASMR1 + ASMR2
                ASMR1 = lbl1.Caption
                ASMR2 = Empty
                SOMA = False
            Else
                If DIVI = True Then
                    VIRG = False
                    If ASMR2 = 0 Then
                        Call MsgBox("Coloque um número que não seja 0")
                    Else
                        lbl1.Caption = ASMR1 / ASMR2
                        ASMR1 = lbl1.Caption
                         ASMR2 = Empty
                    End If
                 Else
                    
                    If SUBT = True Then
                        VIRG = False
                        lbl1.Caption = ASMR1 - ASMR2
                        ASMR1 = lbl1.Caption
                        ASMR2 = Empty
                        SUBT = False
                    End If
                    If MULT = True Then
                        VIRG = False
                        lbl1.Caption = ASMR1 * ASMR2
                        ASMR1 = lbl1.Caption
                        ASMR2 = Empty
                    End If
                    If POTE = True Then
                        VIRG = False
                        lbl1.Caption = ASMR1 ^ ASMR2
                        ASMR1 = lbl1.Caption
                        ASMR2 = Empty
                    End If
                    If RAIZREAL = True Then
                        VIRG = False
                        lbl1.Caption = Math.Sqr(ASMR1)
                        ASMR1 = lbl1.Caption
                        ASMR2 = Empty
                    End If
                    If RAIZ = True Then
                        VIRG = False
                        lbl1.Caption = ASMR1 ^ (1 / ASMR2)
                        ASMR1 = lbl1.Caption
                        ASMR2 = Empty
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdM_Click()
If lbl1.Caption = Empty Then
    Call MsgBox("Insira um número")
Else
    VIRG = False
    If SUBT = False Then
    ASMR1 = lbl1.Caption
    lbl1.Caption = Empty
    SOMA = False
    SUBT = True
    DIVI = False
    SUBTON = False
    POTE = False
    SOMAON = False
    RAIZ = False
    MULT = False
End If
    If SUBT = True Then
        If SUBTON = False Then
            ASMR2 = lbl1.Caption
            ASMR3 = ASMR1 - ASMR2
            lbl1.Caption = Empty
            ASMR2 = Empty
            ASMR1 = ASMR3
            ASMR3 = Empty
            SUBTON = True
            SOMA = False
            SUBT = True
            DIVI = False
            SOMAON = False
            POTE = False
            RAIZ = False
            MULT = False
        Else
        If SUBTON = True Then
            ASMR3 = ASMR1 - CCur(lbl1.Caption)
            ASMR1 = ASMR3
            ASMR3 = Empty
            ASMR2 = Empty
            lbl1.Caption = Empty
            SOMAON = False
            SOMA = False
            SUBT = True
            DIVI = False
            SUBTON = False
            POTE = False
            RAIZ = False
            MULT = False
        End If
    End If
End If
End If
End Sub

Private Sub cmdMp_Click()
If lbl1.Caption = Empty Then
    Call MsgBox("Insira um número")
Else
    VIRG = False
    ASMR1 = CCur(lbl1.Caption)
    lbl1.Caption = Empty
    SOMA = False
    SUBT = False
    SUBTON = False
    SOMATON = False
    DIVI = False
    POTE = False
    RAIZ = False
    MULT = True
End If
End Sub

Private Sub cmdP_Click()
If lbl1.Caption = Empty Then
Call MsgBox("Insira um número")
Else
    VIRG = False
    ASMR1 = CCur(lbl1.Caption)
    lbl1.Caption = Empty
    SOMA = False
    SUBT = False
    DIVI = False
    POTE = True
    RAIZ = False
    MULT = False
    SOMAON = False
    SUBT = False
End If
End Sub

Private Sub cmdPi_Click()
If VIRG = False Then
    lbl1.Caption = lbl1.Caption & "3,1415"
    VIRG = True
End If
End Sub

Private Sub cmdR_Click()
If lbl1.Caption = Empty Then
    Call MsgBox("Insira um número")
Else
    VIRG = False
    ASMR1 = lbl1.Caption
    lbl1.Caption = Empty
    SOMA = False
    SUBT = False
    DIVI = False
    POTE = False
    RAIZ = True
    MULT = False
    SOMAON = False
    SUBT = False
End If
End Sub
Private Sub cmdS_Click()
If lbl1.Caption = Empty Then
    Call MsgBox("Insira um número")
Else
    VIRG = False
    If SOMA = False Then
    ASMR1 = lbl1.Caption
    lbl1.Caption = Empty
    SOMA = True
    SUBT = False
    DIVI = False
    SUBTON = False
    POTE = False
    SOMAON = False
    RAIZ = False
    MULT = False
Else
    If SOMA = True Then
        If SOMAON = False Then
            ASMR2 = lbl1.Caption
            ASMR3 = ASMR1 + ASMR2
            lbl1.Caption = Empty
            ASMR2 = Empty
            ASMR1 = ASMR3
            ASMR3 = Empty
            SOMAON = True
            SOMA = True
            SUBT = False
            DIVI = False
            SUBTON = False
            POTE = False
            RAIZ = False
            MULT = False
        Else
        If SOMAON = True Then
            ASMR3 = CCur(lbl1.Caption) + ASMR1
            ASMR1 = ASMR3
            ASMR3 = Empty
            ASMR2 = Empty
            lbl1.Caption = Empty
            SOMAON = False
            SOMA = True
            SUBT = False
            DIVI = False
            SUBTON = False
            POTE = False
            RAIZ = False
            MULT = False
        End If
    End If
End If
End If
End If
End Sub

Private Sub cmdV_Click()
If VIRG = False Then
    If lbl1.Caption <> Empty Then
        lbl1.Caption = lbl1.Caption & ","
        VIRG = True
    End If
End If
End Sub
