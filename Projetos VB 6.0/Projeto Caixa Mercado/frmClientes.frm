VERSION 5.00
Begin VB.Form frmClientes 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   9075
   ClientLeft      =   10530
   ClientTop       =   4785
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   7800
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Enabled         =   0   'False
      Height          =   510
      Left            =   4770
      TabIndex        =   16
      Top             =   7320
      Width           =   1065
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   510
      Left            =   3600
      TabIndex        =   15
      Top             =   7320
      Width           =   1065
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Enabled         =   0   'False
      Height          =   510
      Left            =   3600
      TabIndex        =   14
      Top             =   6690
      Width           =   1065
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   510
      Left            =   4770
      TabIndex        =   13
      Top             =   6690
      Width           =   1065
   End
   Begin VB.ListBox lstClientes 
      Height          =   2595
      ItemData        =   "frmClientes.frx":0000
      Left            =   555
      List            =   "frmClientes.frx":0002
      TabIndex        =   12
      Top             =   6300
      Width           =   2085
   End
   Begin VB.Frame FrameDad 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Caption         =   "Dados"
      Enabled         =   0   'False
      Height          =   1995
      Left            =   990
      TabIndex        =   2
      Top             =   2820
      Width           =   6540
      Begin VB.ComboBox cmbEstado 
         Height          =   315
         Left            =   3015
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1320
         Width           =   2280
      End
      Begin VB.TextBox txtCidade 
         Height          =   285
         Left            =   255
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1335
         Width           =   1410
      End
      Begin VB.TextBox txtEndereco 
         Height          =   285
         Left            =   3015
         MaxLength       =   100
         TabIndex        =   5
         Top             =   465
         Width           =   2280
      End
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   240
         MaxLength       =   50
         TabIndex        =   3
         Top             =   465
         Width           =   1410
      End
      Begin VB.Label lblEstado 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   3015
         TabIndex        =   9
         Top             =   975
         Width           =   915
      End
      Begin VB.Label lblCidade 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade"
         ForeColor       =   &H00808000&
         Height          =   225
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço"
         ForeColor       =   &H00808000&
         Height          =   180
         Left            =   3015
         TabIndex        =   6
         Top             =   135
         Width           =   1140
      End
      Begin VB.Label lblNome 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   360
         TabIndex        =   4
         Top             =   120
         Width           =   1170
      End
   End
   Begin VB.Frame FrameCod 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Caption         =   "Código"
      Height          =   1155
      Left            =   1290
      TabIndex        =   0
      Top             =   1470
      Width           =   3180
      Begin VB.TextBox txtCodigo 
         Height          =   390
         Left            =   255
         MaxLength       =   10
         TabIndex        =   1
         Top             =   600
         Width           =   2880
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         ForeColor       =   &H00808000&
         Height          =   375
         Left            =   1155
         TabIndex        =   17
         Top             =   150
         Width           =   1095
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Clientes"
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
      Left            =   4905
      TabIndex        =   19
      Top             =   390
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Clientes"
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
      Left            =   4935
      TabIndex        =   18
      Top             =   390
      Width           =   2895
   End
   Begin VB.Label lblCientes 
      BackStyle       =   0  'Transparent
      Caption         =   "Clientes Cadastrados"
      ForeColor       =   &H00808000&
      Height          =   345
      Left            =   630
      TabIndex        =   11
      Top             =   5535
      Width           =   2175
   End
   Begin VB.Menu mnuVlt 
      Caption         =   "Voltar"
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdAlterar_Click()
If IsNumeric(txtNome.Text) = True Then
    MsgBox ("Preencha o nome corretamente"), vbExclamation, "Erro no cadastro"
ElseIf txtNome.Text = "" Then
    MsgBox ("Preencha o nome"), vbExclamation, "Erro no cadastro"
ElseIf txtEndereco.Text = "" Then
    MsgBox ("Preencha o endereço"), vbExclamation, "Erro no cadastro"
ElseIf IsNumeric(txtCidade.Text) = True Then
    MsgBox ("Preencha a cidade corretamente"), vbExclamation, "Erro no cadastro"
ElseIf cmbEstado.Text = "" Then
    MsgBox ("Preencha estado"), vbExclamation, "Erro no cadastro"
ElseIf txtCidade.Text = "" Then
    MsgBox ("Preencha a cidade"), vbExclamation, "Erro no cadastro"
Else
    MsgBox ("Alterado com sucesso!"), vbInformation, "Aviso do cadastro"
    alt = "UPDATE clientes Set Nome ='" & txtNome.Text & "', endereco = '" & txtEndereco.Text & "', cidade = '" & txtCidade.Text & "', estado = '" & cmbEstado.Text & "' where codigo = '" & txtCodigo.Text & "'"
    Conexao.Execute alt
    limp frmClientes
    estado frmClientes
    cmdIncluir.Enabled = False
    cmdExcluir.Enabled = False
    cmdAlterar.Enabled = False
    txtCodigo.Enabled = True
    FrameDad.Enabled = False
    FrameCod.Enabled = True
    carregarlist "clientes", lstClientes, "codigo", "nome"
End If
End Sub

Private Sub cmdExcluir_Click()
DEL = "Delete from clientes where codigo ='" & txtCodigo.Text & "'"
Conexao.Execute (DEL)
limp frmClientes
carregarlist "clientes", lstClientes, "codigo", "nome"
estado frmClientes
cmdIncluir.Enabled = False
cmdExcluir.Enabled = False
cmdAlterar.Enabled = False
txtCodigo.Enabled = True
FrameCod.Enabled = True
FrameDad.Enabled = False
End Sub

Private Sub cmdIncluir_Click()
If IsNumeric(txtNome.Text) = True Then
    MsgBox ("Preencha o nome corretamente"), vbExclamation, "Erro no cadastro"
ElseIf txtNome.Text = "" Then
    MsgBox ("Preencha o nome"), vbExclamation, "Erro no cadastro"
ElseIf txtEndereco.Text = "" Then
    MsgBox ("Preencha o endereço"), vbExclamation, "Erro no cadastro"
ElseIf IsNumeric(txtCidade.Text) = True Then
    MsgBox ("Preencha a cidade corretamente"), vbExclamation, "Erro no cadastro"
ElseIf cmbEstado.Text = "" Then
    MsgBox ("Selecione estado"), vbExclamation, "Erro no cadastro"
ElseIf txtCidade.Text = "" Then
    MsgBox ("Preencha a cidade"), vbExclamation, "Erro no cadastro"
Else
    INS = "Insert into Clientes(Codigo, Nome, Endereco, Cidade, Estado) Values('" & txtCodigo.Text & "','" & txtNome.Text & "','" & txtEndereco.Text & "','" & txtCidade.Text & "','" & cmbEstado.Text & "')"
    Conexao.Execute (INS)
    MsgBox ("Incluído com sucesso!"), vbInformation, "Aviso do cadastro"
    limp frmClientes
    estado frmClientes
    cmdIncluir.Enabled = False
    cmdExcluir.Enabled = False
    cmdAlterar.Enabled = False
    txtCodigo.Enabled = True
    FrameDad.Enabled = False
    carregarlist "clientes", lstClientes, "codigo", "nome"
End If
End Sub

Private Sub cmdLimpar_Click()
limp frmClientes
cmdIncluir.Enabled = False
cmdExcluir.Enabled = False
cmdAlterar.Enabled = False
FrameDad.Enabled = False
FrameCod.Enabled = True
estado frmClientes
txtCodigo.Enabled = True
carregarlist "clientes", lstClientes, "codigo", "nome"
End Sub



Private Sub Form_Load()
estado frmClientes
carregarlist "clientes", lstClientes, "codigo", "nome"
End Sub




Private Sub lstClientes_Click()
Dim codigo As String
Dim posicao As String
posicao = InStr(1, lstClientes, "-")
codigo = Left(lstClientes, posicao - 2)
sql = "Select  * from clientes where codigo = '" & codigo & "'"
Set tabela = Conexao.Execute(sql)
If tabela.EOF Then
    MsgBox "nada cadastrado"
Else
    txtCidade.Text = tabela("cidade")
    cmbEstado = tabela("estado")
    txtEndereco.Text = tabela("endereco")
    txtCodigo.Text = tabela("codigo")
    txtNome.Text = tabela("nome")
    FrameCod.Enabled = False
    FrameDad.Enabled = True
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = True
    cmdExcluir.Enabled = True
End If
End Sub

Private Sub mnuVlt_Click()
Unload frmClientes
End Sub


Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And IsNumeric(txtCodigo.Text) = True Then
    If VerCod("clientes", txtCodigo.Text, "codigo") = False Then
     MsgBox ("Código já existente")
    Else
        cmdIncluir.Enabled = True
        txtCodigo.Enabled = False
        FrameDad.Enabled = True
    End If
End If
End Sub

