VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compras"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11355
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11355
   Begin VB.ComboBox cmbProduto1 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2640
      Width           =   2895
   End
   Begin VB.ComboBox cmbProduto2 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3240
      Width           =   2895
   End
   Begin VB.ComboBox cmbProduto3 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3840
      Width           =   2895
   End
   Begin VB.ComboBox cmbProduto4 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4440
      Width           =   2895
   End
   Begin VB.ComboBox cmbProduto5 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5040
      Width           =   2895
   End
   Begin VB.TextBox txtQtd1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      TabIndex        =   4
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox txtQtd5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      TabIndex        =   3
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox txtQtd4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      TabIndex        =   2
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox txtQtd3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      TabIndex        =   1
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox txtQtd2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      TabIndex        =   0
      Top             =   3240
      Width           =   615
   End
   Begin VB.Line Line12 
      X1              =   8040
      X2              =   8040
      Y1              =   5520
      Y2              =   1800
   End
   Begin VB.Line Line11 
      X1              =   6120
      X2              =   6120
      Y1              =   5520
      Y2              =   1800
   End
   Begin VB.Line Line10 
      X1              =   4440
      X2              =   4440
      Y1              =   5520
      Y2              =   1800
   End
   Begin VB.Line Line9 
      X1              =   1200
      X2              =   1200
      Y1              =   1800
      Y2              =   5520
   End
   Begin VB.Line Line8 
      X1              =   9960
      X2              =   1200
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line7 
      X1              =   9960
      X2              =   1200
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line6 
      X1              =   9960
      X2              =   1200
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line5 
      X1              =   9960
      X2              =   1200
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line4 
      X1              =   9960
      X2              =   1320
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line3 
      X1              =   9960
      X2              =   1200
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      X1              =   9960
      X2              =   9960
      Y1              =   1800
      Y2              =   5520
   End
   Begin VB.Line Line1 
      X1              =   1200
      X2              =   9960
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label lbl4 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total"
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
      Left            =   8280
      TabIndex        =   24
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade"
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
      Left            =   6360
      TabIndex        =   23
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
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
      Left            =   4920
      TabIndex        =   22
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Produto"
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
      Left            =   2280
      TabIndex        =   21
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblValor1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblValor2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4680
      TabIndex        =   19
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblValor3 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4680
      TabIndex        =   18
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblValor4 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4680
      TabIndex        =   17
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label lblValor5 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label lblValorTotal1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8280
      TabIndex        =   15
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblValorTotal2 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8280
      TabIndex        =   14
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblValorTotal3 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8280
      TabIndex        =   13
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblValorTotal4 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8280
      TabIndex        =   12
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label lblValorTotal5 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8280
      TabIndex        =   11
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label lblCompras 
      BackStyle       =   0  'Transparent
      Caption         =   "Compras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   10
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbProduto2_Click()
txtQtd2.Enabled = True
If cmbProduto2.Text = "Ervilha" Then
    lblValor2.Caption = "1,00"
    If IsNumeric(txtQtd2.Text) = True Then
        If CCur(txtQtd2.Text) <= 50 Then
            lblValorTotal2.Caption = CCur(txtQtd2.Text) * lblValor2.Caption
        End If
    End If
ElseIf cmbProduto2.Text = "Abacaxi" Then
    lblValor2.Caption = "3,80"
    If IsNumeric(txtQtd2.Text) = True Then
        If CCur(txtQtd2.Text) <= 50 Then
            lblValorTotal2.Caption = CCur(txtQtd2.Text) * lblValor2.Caption
        End If
    End If
ElseIf cmbProduto2.Text = "Melão" Then
    lblValor2.Caption = "4,50"
    If IsNumeric(txtQtd2.Text) = True Then
        If CCur(txtQtd2.Text) <= 50 Then
            lblValorTotal2.Caption = CCur(txtQtd2.Text) * lblValor2.Caption
        End If
    End If
ElseIf cmbProduto2.Text = "Milho" Then
    lblValor2.Caption = "1,20"
    If IsNumeric(txtQtd2.Text) = True Then
        If CCur(txtQtd2.Text) <= 50 Then
            lblValorTotal2.Caption = CCur(txtQtd2.Text) * lblValor2.Caption
        End If
    End If
ElseIf cmbProduto2.Text = "Ovo" Then
    lblValor2.Caption = "2,30"
    If IsNumeric(txtQtd2.Text) = True Then
        If CCur(txtQtd2.Text) <= 50 Then
            lblValorTotal2.Caption = CCur(txtQtd2.Text) * lblValor2.Caption
        End If
    End If
End If
End Sub
Private Sub cmbProduto1_Click()
txtQtd1.Enabled = True
If cmbProduto1.Text = "Ervilha" Then
    lblValor1.Caption = "1,00"
    If IsNumeric(txtQtd1.Text) = True Then
        If CCur(txtQtd1.Text) <= 50 Then
            lblValorTotal1.Caption = CCur(txtQtd1.Text) * lblValor1.Caption
        End If
    End If
ElseIf cmbProduto1.Text = "Abacaxi" Then
    lblValor1.Caption = "3,80"
    If IsNumeric(txtQtd1.Text) = True Then
        If CCur(txtQtd1.Text) <= 50 Then
            lblValorTotal1.Caption = Format(CCur(txtQtd1.Text) * lblValor1.Caption, "R$0")
        End If
    End If
ElseIf cmbProduto1.Text = "Melão" Then
    lblValor1.Caption = "4,50"
    If IsNumeric(txtQtd1.Text) = True Then
        If CCur(txtQtd1.Text) <= 50 Then
            lblValorTotal1.Caption = CCur(txtQtd1.Text) * lblValor1.Caption
        End If
    End If
ElseIf cmbProduto1.Text = "Milho" Then
    lblValor1.Caption = "1,20"
    If IsNumeric(txtQtd1.Text) = True Then
        If CCur(txtQtd1.Text) <= 50 Then
            lblValorTotal1.Caption = CCur(txtQtd1.Text) * lblValor1.Caption
        End If
    End If
ElseIf cmbProduto1.Text = "Ovo" Then
    lblValor1.Caption = "2,30"
    If IsNumeric(txtQtd1.Text) = True Then
        If CCur(txtQtd1.Text) <= 50 Then
            lblValorTotal1.Caption = CCur(txtQtd1.Text) * lblValor1.Caption
        End If
    End If
End If
End Sub
Private Sub cmbProduto5_Click()
txtQtd5.Enabled = True
If cmbProduto5.Text = "Ervilha" Then
    lblValor5.Caption = "1,00"
    If IsNumeric(txtQtd5.Text) = True Then
        If CCur(txtQtd5.Text) <= 50 Then
            lblValorTotal5.Caption = CCur(txtQtd5.Text) * lblValor5.Caption
        End If
    End If
ElseIf cmbProduto5.Text = "Abacaxi" Then
    lblValor5.Caption = "3,80"
    If IsNumeric(txtQtd5.Text) = True Then
        If CCur(txtQtd5.Text) <= 50 Then
            lblValorTotal5.Caption = CCur(txtQtd5.Text) * lblValor5.Caption
        End If
    End If
ElseIf cmbProduto5.Text = "Melão" Then
    lblValor5.Caption = "4,50"
    If IsNumeric(txtQtd5.Text) = True Then
        If CCur(txtQtd5.Text) <= 50 Then
            lblValorTotal5.Caption = CCur(txtQtd5.Text) * lblValor5.Caption
        End If
    End If
ElseIf cmbProduto5.Text = "Milho" Then
    lblValor5.Caption = "1,20"
    If IsNumeric(txtQtd5.Text) = True Then
        If CCur(txtQtd5.Text) <= 50 Then
            lblValorTotal5.Caption = CCur(txtQtd5.Text) * lblValor5.Caption
        End If
    End If
ElseIf cmbProduto5.Text = "Ovo" Then
    lblValor5.Caption = "2,30"
    If IsNumeric(txtQtd5.Text) = True Then
        If CCur(txtQtd5.Text) <= 50 Then
            lblValorTotal5.Caption = CCur(txtQtd5.Text) * lblValor5.Caption
        End If
    End If
End If
End Sub
Private Sub cmbProduto4_Click()
txtQtd4.Enabled = True
If cmbProduto4.Text = "Ervilha" Then
    lblValor4.Caption = "1,00"
    If IsNumeric(txtQtd4.Text) = True Then
        If CCur(txtQtd4.Text) <= 50 Then
            lblValorTotal4.Caption = CCur(txtQtd4.Text) * lblValor4.Caption
        End If
    End If
ElseIf cmbProduto4.Text = "Abacaxi" Then
    lblValor4.Caption = "3,80"
    If IsNumeric(txtQtd4.Text) = True Then
        If CCur(txtQtd4.Text) <= 50 Then
            lblValorTotal4.Caption = CCur(txtQtd4.Text) * lblValor4.Caption
        End If
    End If
ElseIf cmbProduto4.Text = "Melão" Then
    lblValor4.Caption = "4,50"
    If IsNumeric(txtQtd4.Text) = True Then
        If CCur(txtQtd4.Text) <= 50 Then
            lblValorTotal4.Caption = CCur(txtQtd4.Text) * lblValor4.Caption
        End If
    End If
ElseIf cmbProduto4.Text = "Milho" Then
    lblValor4.Caption = "1,20"
    If IsNumeric(txtQtd4.Text) = True Then
        If CCur(txtQtd4.Text) <= 50 Then
            lblValorTotal4.Caption = CCur(txtQtd4.Text) * lblValor4.Caption
        End If
    End If
ElseIf cmbProduto4.Text = "Ovo" Then
    lblValor4.Caption = "2,30"
    If IsNumeric(txtQtd4.Text) = True Then
        If CCur(txtQtd4.Text) <= 50 Then
            lblValorTotal4.Caption = CCur(txtQtd4.Text) * lblValor4.Caption
        End If
    End If
End If
End Sub
Private Sub cmbProduto3_Click()
txtQtd3.Enabled = True
If cmbProduto3.Text = "Ervilha" Then
    lblValor3.Caption = "1,00"
    If IsNumeric(txtQtd3.Text) = True Then
        If CCur(txtQtd3.Text) <= 50 Then
            lblValorTotal3.Caption = CCur(txtQtd3.Text) * lblValor3.Caption
        End If
    End If
ElseIf cmbProduto3.Text = "Abacaxi" Then
    lblValor3.Caption = "3,80"
    If IsNumeric(txtQtd3.Text) = True Then
        If CCur(txtQtd3.Text) <= 50 Then
            lblValorTotal3.Caption = CCur(txtQtd3.Text) * lblValor3.Caption
        End If
    End If
ElseIf cmbProduto3.Text = "Melão" Then
    lblValor3.Caption = "4,50"
    If IsNumeric(txtQtd3.Text) = True Then
        If CCur(txtQtd3.Text) <= 50 Then
            lblValorTotal3.Caption = CCur(txtQtd3.Text) * lblValor3.Caption
        End If
    End If
ElseIf cmbProduto3.Text = "Milho" Then
    lblValor3.Caption = "1,20"
    If IsNumeric(txtQtd3.Text) = True Then
        If CCur(txtQtd3.Text) <= 50 Then
            lblValorTotal3.Caption = CCur(txtQtd3.Text) * lblValor3.Caption
        End If
    End If
ElseIf cmbProduto3.Text = "Ovo" Then
    lblValor3.Caption = "2,30"
    If IsNumeric(txtQtd3.Text) = True Then
        If CCur(txtQtd3.Text) <= 50 Then
            lblValorTotal3.Caption = CCur(txtQtd3.Text) * lblValor3.Caption
        End If
    End If
End If
End Sub

Private Sub cmbProduto1_GotFocus()
cmbProduto1.BackColor = &H80FFFF
End Sub

Private Sub cmbProduto1_LostFocus()
cmbProduto1.BackColor = &HFFFFFF
End Sub
Private Sub cmbProduto2_GotFocus()
cmbProduto2.BackColor = &H80FFFF
End Sub

Private Sub cmbProduto2_LostFocus()
cmbProduto2.BackColor = &HFFFFFF
End Sub
Private Sub cmbProduto3_GotFocus()
cmbProduto3.BackColor = &H80FFFF
End Sub

Private Sub cmbProduto3_LostFocus()
cmbProduto3.BackColor = &HFFFFFF
End Sub
Private Sub cmbProduto4_GotFocus()
cmbProduto4.BackColor = &H80FFFF
End Sub

Private Sub cmbProduto4_LostFocus()
cmbProduto4.BackColor = &HFFFFFF
End Sub
Private Sub cmbProduto5_GotFocus()
cmbProduto5.BackColor = &H80FFFF
End Sub
Private Sub cmbProduto5_LostFocus()
cmbProduto5.BackColor = &HFFFFFF
End Sub
Private Sub Form_Load()
cmbProduto1.AddItem "Abacaxi"
cmbProduto1.AddItem "Ervilha"
cmbProduto1.AddItem "Melão"
cmbProduto1.AddItem "Ovo"
cmbProduto1.AddItem "Milho"
cmbProduto2.AddItem "Abacaxi"
cmbProduto2.AddItem "Ervilha"
cmbProduto2.AddItem "Melão"
cmbProduto2.AddItem "Ovo"
cmbProduto2.AddItem "Milho"
cmbProduto3.AddItem "Abacaxi"
cmbProduto3.AddItem "Ervilha"
cmbProduto3.AddItem "Melão"
cmbProduto3.AddItem "Ovo"
cmbProduto3.AddItem "Milho"
cmbProduto4.AddItem "Abacaxi"
cmbProduto4.AddItem "Ervilha"
cmbProduto4.AddItem "Melão"
cmbProduto4.AddItem "Ovo"
cmbProduto4.AddItem "Milho"
cmbProduto5.AddItem "Abacaxi"
cmbProduto5.AddItem "Ervilha"
cmbProduto5.AddItem "Melão"
cmbProduto5.AddItem "Ovo"
cmbProduto5.AddItem "Milho"
End Sub

Private Sub txtQtd1_Change()
If IsNumeric(txtQtd1.Text) = True And (cmbProduto1.Text = "") = False Then
    lblValorTotal1.Caption = CCur(txtQtd1.Text) * lblValor1.Caption
ElseIf IsNumeric(txtQtd1.Text) = False And (txtQtd1.Text = "") = False Then
    MsgBox ("Coloque apenas números")
    txtQtd1.Text = ""
    lblValorTotal1.Caption = ""
End If
End Sub
Private Sub txtQtd2_Change()
If IsNumeric(txtQtd2.Text) = True And (cmbProduto2.Text = "") = False Then
    lblValorTotal2.Caption = CCur(txtQtd2.Text) * lblValor2.Caption
ElseIf IsNumeric(txtQtd2.Text) = False And (txtQtd2.Text = "") = False Then
    MsgBox ("Coloque apenas números")
    txtQtd2.Text = ""
    lblValorTotal2.Caption = ""
End If
End Sub
Private Sub txtQtd3_Change()
If IsNumeric(txtQtd3.Text) = True And (cmbProduto3.Text = "") = False Then
    lblValorTotal3.Caption = CCur(txtQtd3.Text) * lblValor3.Caption
ElseIf IsNumeric(txtQtd3.Text) = False And (txtQtd3.Text = "") = False Then
    MsgBox ("Coloque apenas números")
    txtQtd3.Text = ""
    lblValorTotal4.Caption = ""
End If
End Sub
Private Sub txtQtd4_Change()
If IsNumeric(txtQtd4.Text) = True And (cmbProduto4.Text = "") = False Then
    lblValorTotal4.Caption = CCur(txtQtd4.Text) * lblValor4.Caption
ElseIf IsNumeric(txtQtd4.Text) = False And (txtQtd4.Text = "") = False Then
    MsgBox ("Coloque apenas números")
    txtQtd4.Text = ""
    lblValorTotal4.Caption = ""
End If
End Sub
Private Sub txtQtd5_Change()
If IsNumeric(txtQtd5.Text) = True And (cmbProduto5.Text = "") = False Then
    lblValorTotal5.Caption = CCur(txtQtd5.Text) * lblValor5.Caption
ElseIf IsNumeric(txtQtd5.Text) = False And (txtQtd5.Text = "") = False Then
    MsgBox ("Coloque apenas números")
    txtQtd5.Text = ""
    lblValorTotal5.Caption = ""
End If
End Sub

Private Sub txtQtd1_GotFocus()
txtQtd1.BackColor = &H80FFFF
End Sub

Private Sub txtQtd1_LostFocus()
txtQtd1.BackColor = &HFFFFFF
End Sub

Private Sub txtQtd2_GotFocus()
txtQtd2.BackColor = &H80FFFF
End Sub

Private Sub txtQtd2_LostFocus()
txtQtd2.BackColor = &HFFFFFF
End Sub
Private Sub txtQtd3_GotFocus()
txtQtd3.BackColor = &H80FFFF
End Sub
Private Sub txtQtd3_LostFocus()
txtQtd3.BackColor = &HFFFFFF
End Sub
Private Sub txtQtd4_GotFocus()
txtQtd4.BackColor = &H80FFFF
End Sub
Private Sub txtQtd4_LostFocus()
txtQtd4.BackColor = &HFFFFFF
End Sub
Private Sub txtQtd5_GotFocus()
txtQtd5.BackColor = &H80FFFF
End Sub
Private Sub txtQtd5_LostFocus()
txtQtd5.BackColor = &HFFFFFF
End Sub


