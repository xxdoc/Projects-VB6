VERSION 5.00
Begin VB.Form frmSetor 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Setores"
   ClientHeight    =   9345
   ClientLeft      =   10485
   ClientTop       =   4785
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   7740
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   510
      Left            =   5850
      TabIndex        =   14
      Top             =   5160
      Width           =   1065
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Enabled         =   0   'False
      Height          =   510
      Left            =   4680
      TabIndex        =   13
      Top             =   5160
      Width           =   1065
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   510
      Left            =   4680
      TabIndex        =   12
      Top             =   5790
      Width           =   1065
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Enabled         =   0   'False
      Height          =   510
      Left            =   5850
      TabIndex        =   11
      Top             =   5790
      Width           =   1065
   End
   Begin VB.ListBox lstSeto 
      Height          =   2400
      Left            =   315
      TabIndex        =   9
      Top             =   6225
      Width           =   3675
   End
   Begin VB.Frame fmDad 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Caption         =   "Dados"
      Height          =   1680
      Left            =   300
      TabIndex        =   5
      Top             =   3525
      Width           =   5970
      Begin VB.TextBox txtSetores 
         Enabled         =   0   'False
         Height          =   315
         Left            =   330
         MaxLength       =   3
         TabIndex        =   6
         Top             =   780
         Width           =   3210
      End
      Begin VB.Label label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Setor"
         ForeColor       =   &H00808000&
         Height          =   390
         Left            =   330
         TabIndex        =   7
         Top             =   225
         Width           =   1800
      End
   End
   Begin VB.Frame fmCode 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Caption         =   "Código"
      Height          =   1725
      Left            =   390
      TabIndex        =   2
      Top             =   1650
      Width           =   3705
      Begin VB.TextBox txtCodigo 
         Height          =   345
         Left            =   255
         MaxLength       =   3
         TabIndex        =   3
         Top             =   1155
         Width           =   2955
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   255
         TabIndex        =   8
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   ","
         ForeColor       =   &H00808000&
         Height          =   390
         Left            =   -15
         TabIndex        =   4
         Top             =   285
         Width           =   2160
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Setores Cadastrados"
      ForeColor       =   &H00808000&
      Height          =   405
      Left            =   330
      TabIndex        =   10
      Top             =   5595
      Width           =   2565
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Setores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   600
      Left            =   3795
      TabIndex        =   1
      Top             =   900
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Setores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   600
      Left            =   3825
      TabIndex        =   0
      Top             =   915
      Width           =   2895
   End
   Begin VB.Menu fmVolt 
      Caption         =   "Voltar"
   End
End
Attribute VB_Name = "frmSetor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAlterar_Click()
If IsNumeric(txtSetores.Text) = True Then
    MsgBox ("Preencha com um nome de setor"), vbExclamation, "Erro no cadastro"
ElseIf txtSetores.Text = "" Then
    MsgBox ("Preencha o setor"), vbExclamation, "Erro no cadastro"
Else
    alt = "UPDATE setor Set Setor ='" & txtSetores.Text & "' where Código = '" & txtCodigo.Text & "'"
    Conexao.Execute (alt)
    limp frmSetor
    cmdIncluir.Enabled = False
    cmdExcluir.Enabled = False
    cmdAlterar.Enabled = False
    fmDad.Enabled = False
    fmCode.Enabled = True
    carregarlist "Setor", lstSeto, "Setor", "Código"
End If
End Sub

Private Sub cmdExcluir_Click()
DEL = "delete from Setor where Código ='" & txtCodigo.Text & "'"
Conexao.Execute (DEL)
carregarlist "Setor", lstSeto, "Setor", "Código"
limp frmSetor
cmdIncluir.Enabled = False
cmdExcluir.Enabled = False
cmdAlterar.Enabled = False
fmCode.Enabled = True
fmDad.Enabled = False
carregarlist "Setor", lstSeto, "Setor", "Código"
End Sub

Private Sub cmdIncluir_Click()
If IsNumeric(txtSetores.Text) = True Then
    MsgBox ("Preencha com um nome de setor"), vbExclamation, "Erro no cadastro"
ElseIf txtSetores.Text = "" Then
    MsgBox ("Preencha o setor"), vbExclamation, "Erro no cadastro"
Else
    INS = "Insert into Setor(Código, Setor) Values('" & txtCodigo.Text & "','" & txtSetores.Text & "')"
    Conexao.Execute (INS)
    limp frmSetor
    cmdIncluir.Enabled = False
    cmdExcluir.Enabled = False
    cmdAlterar.Enabled = False
    fmCode.Enabled = True
    fmDad.Enabled = False
    MsgBox ("Incluído com sucesso!"), vbInformation, "Aviso do cadastro"
    carregarlist "Setor", lstSeto, "Setor", "Código"
End If
End Sub

Private Sub cmdLimpar_Click()
limp frmSetor
cmdIncluir.Enabled = False
cmdExcluir.Enabled = False
cmdAlterar.Enabled = False
fmDad.Enabled = False
fmCode.Enabled = True
carregarlist "Setor", lstSeto, "Setor", "Código"

End Sub

Private Sub fmVolt_Click()
Unload frmSetor
End Sub

Private Sub Form_Load()
carregarlist "Setor", lstSeto, "Setor", "Código"
End Sub


Private Sub lstSeto_Click()
Dim um, dois As String
dois = InStr(1, lstSeto, "-")
um = Left(lstSeto, dois - 2)
SQL = "Select * from Setor where Setor = '" & um & "'"
Set tabela = Conexao.Execute(SQL)
If tabela.EOF Then
    MsgBox "nada cadastrado"
Else
    txtCodigo.Text = tabela("Código")
    txtSetores.Text = tabela("Setor")
    fmCode.Enabled = False
    fmDad.Enabled = True
    txtSetores.Enabled = True
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = True
    cmdExcluir.Enabled = True
End If
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Len(txtCodigo.Text) <> 0 And IsNumeric(txtCodigo.Text) Then
    
    If VerCod("Setor", txtCodigo.Text, "Código") = False Then
        MsgBox ("Código já existente")
    Else
        fmCode.Enabled = False
        fmDad.Enabled = True
        txtSetores.Enabled = True
        cmdIncluir.Enabled = True
    End If
End If
End Sub
