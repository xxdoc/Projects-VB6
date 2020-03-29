VERSION 5.00
Begin VB.Form frmFornecedores 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fornecedores"
   ClientHeight    =   9045
   ClientLeft      =   10680
   ClientTop       =   5175
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   7740
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   360
      Left            =   5235
      TabIndex        =   18
      Top             =   6405
      Width           =   900
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   435
      Left            =   5250
      TabIndex        =   17
      Top             =   5925
      Width           =   900
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Height          =   360
      Left            =   4170
      TabIndex        =   16
      Top             =   6405
      Width           =   975
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   435
      Left            =   4185
      TabIndex        =   15
      Top             =   5925
      Width           =   945
   End
   Begin VB.ListBox lstFornecedores 
      Height          =   2400
      Left            =   465
      TabIndex        =   14
      Top             =   5670
      Width           =   2910
   End
   Begin VB.Frame fmDados 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Caption         =   "Dados"
      Height          =   3135
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   7335
      Begin VB.ComboBox cmbEstado 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1665
         Width           =   2400
      End
      Begin VB.TextBox txtemail 
         Height          =   285
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   11
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtcidade 
         Height          =   285
         Left            =   360
         MaxLength       =   20
         TabIndex        =   10
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtcontato 
         Height          =   285
         Left            =   3480
         MaxLength       =   13
         TabIndex        =   8
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtendereco 
         Height          =   285
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   7
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   360
         MaxLength       =   50
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Contato"
         Height          =   375
         Left            =   3480
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço"
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblNome 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   225
         Width           =   855
      End
   End
   Begin VB.Frame fmCNPJ 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Caption         =   "CNPJ"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   930
      Width           =   3255
      Begin VB.TextBox txtCNPJ 
         Height          =   375
         Left            =   360
         MaxLength       =   14
         TabIndex        =   1
         Top             =   345
         Width           =   3015
      End
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "CNPJ"
      Height          =   210
      Left            =   750
      TabIndex        =   21
      Top             =   645
      Width           =   1275
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedores"
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
      Left            =   4260
      TabIndex        =   20
      Top             =   735
      Width           =   3315
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedores"
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
      Left            =   4290
      TabIndex        =   19
      Top             =   750
      Width           =   3105
   End
   Begin VB.Menu Voltar 
      Caption         =   "Voltar"
   End
End
Attribute VB_Name = "frmFornecedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAlterar_Click()
If txtNome.Text = "" Or txtendereco.Text = "" Or txtcidade.Text = "" Or cmbEstado.Text = "" Or txtemail.Text = "" Or txtcontato.Text = "" Or IsNumeric(txtcontato.Text) = False Or IsNumeric(txtNome.Text) = True Or IsNumeric(txtcidade.Text) = True Then
    MsgBox "Preencha os campos corretamente", vbCritical, "Aviso do sistema"
Else
    SQL = "update fornecedores SET Nome = '" & txtNome.Text & "', Endereco = '" & txtendereco.Text & "', Cidade = '" & txtcidade.Text & "', Email = '" & txtemail.Text & "', Contato = '" & txtcontato.Text & "', Estado = '" & cmbEstado.Text & "' where CNPJ = '" & txtCNPJ.Text & "'"
    Conexao.Execute (SQL)
    limp frmFornecedores
    fmCNPJ.Enabled = True
    fmDados.Enabled = False
    cmdIncluir.Enabled = False
    cmdExcluir.Enabled = False
    cmdAlterar.Enabled = False
    estado frmFornecedores
    SQL = "select * from fornecedores"
    Set tabeladinamica = Conexao.Execute(SQL)
    If tabeladinamica.EOF Then

    Else
        Do While Not tabeladinamica.EOF
            lstFornecedores.AddItem tabeladinamica("CNPJ") & "-" & tabeladinamica("Nome")
            tabeladinamica.MoveNext
        Loop
    End If
End If
End Sub

Private Sub cmdExcluir_Click()
SQL = "delete from fornecedores where CNPJ = '" & txtCNPJ.Text & "'"
Conexao.Execute (SQL)
limp frmFornecedores
fmCNPJ.Enabled = True
fmDados.Enabled = False
cmdIncluir.Enabled = False
cmdExcluir.Enabled = False
cmdAlterar.Enabled = False
estado frmFornecedores
SQL = "select * from fornecedores"
Set tabeladinamica = Conexao.Execute(SQL)
If tabeladinamica.EOF Then
Else
    lstFornecedores.Clear
        Do While Not tabeladinamica.EOF
            lstFornecedores.AddItem tabeladinamica("CNPJ") & "-" & tabeladinamica("Nome")
            tabeladinamica.MoveNext
        Loop
End If
End Sub


Private Sub cmdIncluir_Click()
If txtNome.Text = "" Or txtendereco.Text = "" Or txtcidade.Text = "" Or cmbEstado.Text = "" Or txtcontato = "" Or txtemail = "" Or IsNumeric(txtcidade.Text) = True Or IsNumeric(txtNome.Text) = True Or IsNumeric(txtcontato.Text) = False Then
    MsgBox "Preencha os campos corretamente", vbCritical, "Aviso do sistema"
Else
    SQL = "Insert into fornecedores (CNPJ, Nome, Endereco, Cidade, Contato, Email, Estado) Values ('" & txtCNPJ.Text & "','" & txtNome.Text & "','" & txtendereco.Text & "','" & txtcidade.Text & "','" & txtcontato.Text & "','" & txtemail.Text & "','" & cmbEstado.Text & "')"
    Conexao.Execute (SQL)
    limp frmFornecedores
    fmCNPJ.Enabled = True
    fmDados.Enabled = False
    cmdIncluir.Enabled = False
    cmdExcluir.Enabled = False
    cmdAlterar.Enabled = False
    cmdLimpar.Enabled = False
    estado frmFornecedores
End If
SQL = "select * from fornecedores"
Set tabeladinamica = Conexao.Execute(SQL)
    If tabeladinamica.EOF Then
    Else
        lstFornecedores.Clear
            Do While Not tabeladinamica.EOF
                lstFornecedores.AddItem tabeladinamica("CNPJ") & "-" & tabeladinamica("Nome")
                    tabeladinamica.MoveNext
            Loop
    End If
End Sub

Private Sub cmdLimpar_Click()
limp frmFornecedores
fmCNPJ.Enabled = True
fmDados.Enabled = False
cmdIncluir.Enabled = False
cmdExcluir.Enabled = False
cmdAlterar.Enabled = False
estado frmFornecedores
End Sub

Private Sub Form_Load()
fmDados.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdIncluir.Enabled = False
cmdLimpar.Enabled = False
SQL = "select * from fornecedores"
Set tabeladinamica = Conexao.Execute(SQL)
    If tabeladinamica.EOF Then
    Else
        lstFornecedores.Clear
        Do While Not tabeladinamica.EOF
            lstFornecedores.AddItem tabeladinamica("CNPJ") & "-" & tabeladinamica("Nome")
            tabeladinamica.MoveNext
        Loop
    End If
estado frmFornecedores
End Sub

Private Sub lstFornecedores_Click()
posicao = InStr(1, lstFornecedores, "-")
cnpj = Left(lstFornecedores, posicao - 1)
SQL = "select * from Fornecedores where CNPJ = '" & (cnpj) & "'"
Set tabela = Conexao.Execute(SQL)
If tabela.EOF Then
Else
    txtCNPJ.Text = tabela("CNPJ")
    txtNome.Text = tabela("Nome")
    txtendereco.Text = tabela("Endereco")
    txtcidade.Text = tabela("Cidade")
    txtemail.Text = tabela("Email")
    txtcontato.Text = tabela("Contato")
    cmbEstado.Text = tabela("Estado")
    fmCNPJ.Enabled = False
    fmDados.Enabled = True
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = True
    cmdExcluir.Enabled = True
End If
End Sub

Private Sub txtCNPJ_Keypress(KeyAscii As Integer)
If KeyAscii = 13 And IsNumeric(txtCNPJ.Text) = True Then
    If VerCod("fornecedores", txtCNPJ.Text, "CNPJ") = False Then
        MsgBox ("CNPJ já existente")
    Else
        fmCNPJ.Enabled = False
        fmDados.Enabled = True
        cmdIncluir.Enabled = True
        cmdAlterar.Enabled = False
        cmdExcluir.Enabled = False
        cmdLimpar.Enabled = True
    End If
End If
End Sub


Private Sub Voltar_Click()
Unload frmFornecedores
End Sub


