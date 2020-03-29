VERSION 5.00
Begin VB.Form frmProdutos 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produtos"
   ClientHeight    =   9075
   ClientLeft      =   10530
   ClientTop       =   4935
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   7800
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   480
      Left            =   2010
      TabIndex        =   22
      Top             =   7455
      Width           =   975
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Enabled         =   0   'False
      Height          =   480
      Left            =   900
      TabIndex        =   21
      Top             =   7455
      Width           =   960
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   480
      Left            =   2010
      TabIndex        =   20
      Top             =   6855
      Width           =   960
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Enabled         =   0   'False
      Height          =   480
      Left            =   900
      TabIndex        =   19
      Top             =   6855
      Width           =   960
   End
   Begin VB.Frame fmProd 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Caption         =   "Produtos"
      Height          =   3795
      Left            =   4020
      TabIndex        =   16
      Top             =   5115
      Width           =   3180
      Begin VB.ListBox lstProd 
         Height          =   2790
         Left            =   480
         TabIndex        =   17
         Top             =   630
         Width           =   2325
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Produtos"
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   525
         TabIndex        =   18
         Top             =   300
         Width           =   1485
      End
   End
   Begin VB.Frame fmDad 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   3750
      Left            =   780
      TabIndex        =   5
      Top             =   2610
      Width           =   5220
      Begin VB.ComboBox cmbForn 
         Height          =   315
         Left            =   285
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3390
         Width           =   1695
      End
      Begin VB.ComboBox cmbSeto 
         Height          =   315
         Left            =   2970
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   810
         Width           =   1935
      End
      Begin VB.TextBox txtValo 
         Height          =   345
         Left            =   3045
         TabIndex        =   13
         Top             =   1920
         Width           =   1635
      End
      Begin VB.TextBox txtQTDE 
         Height          =   330
         Left            =   285
         TabIndex        =   12
         Top             =   1920
         Width           =   1755
      End
      Begin VB.TextBox txtDesc 
         Height          =   285
         Left            =   285
         MaxLength       =   30
         TabIndex        =   11
         Top             =   810
         Width           =   1635
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   285
         TabIndex        =   10
         Top             =   2790
         Width           =   1485
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor (R$)"
         ForeColor       =   &H00808000&
         Height          =   390
         Left            =   3045
         TabIndex        =   9
         Top             =   1500
         Width           =   2145
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Setor"
         ForeColor       =   &H00808000&
         Height          =   360
         Left            =   3045
         TabIndex        =   8
         Top             =   465
         Width           =   2040
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade em estoque"
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   285
         TabIndex        =   7
         Top             =   1500
         Width           =   1845
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   285
         TabIndex        =   6
         Top             =   465
         Width           =   1605
      End
   End
   Begin VB.Frame fmCod 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1530
      Left            =   810
      TabIndex        =   2
      Top             =   600
      Width           =   3330
      Begin VB.TextBox txtCod 
         Height          =   315
         Left            =   210
         MaxLength       =   13
         TabIndex        =   3
         Top             =   1050
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   285
         TabIndex        =   4
         Top             =   645
         Width           =   1725
      End
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Produtos"
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
      Left            =   4965
      TabIndex        =   1
      Top             =   705
      Width           =   3390
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Produtos"
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
      Left            =   4995
      TabIndex        =   0
      Top             =   720
      Width           =   3285
   End
   Begin VB.Menu mnuVolt 
      Caption         =   "Voltar"
   End
End
Attribute VB_Name = "frmProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub cmdAlterar_Click()
If txtDesc.Text = "" Then
    MsgBox ("Preencha a descrição")
ElseIf IsNumeric(txtDesc.Text) Then
    MsgBox ("Preencha a descrição corretamente"), vbExclamation, "Erro no cadastro"
ElseIf cmbSeto.Text = "" Then
    MsgBox ("Selecione o setor"), vbExclamation, "Erro no cadastro"
ElseIf txtQTDE.Text = "" Then
    MsgBox ("Preencha a quantidade em estoque"), vbExclamation, "Erro no cadastro"
ElseIf Not IsNumeric(txtQTDE.Text) Then
    MsgBox ("Insira apenas números na quantidade em estoque"), vbExclamation, "Erro no cadastro"
ElseIf txtValo.Text = "" Then
    MsgBox ("Insira o valor do produto"), vbExclamation, "Erro no cadastro"
ElseIf Not IsNumeric(txtValo.Text) Then
    MsgBox ("Insira apenas números no valor do produto"), vbExclamation, "Erro no cadastro"
ElseIf cmbForn.Text = "" Then
    MsgBox ("Selecione o fornecedor"), vbExclamation, "Erro no cadastro"
Else
    SQL = "UPDATE Produtos Set Descricao ='" & txtDesc.Text & "', QTDEstoque ='" & txtQTDE.Text & "', Setor ='" & cmbSeto.Text & "', Fornecedor ='" & cmbForn.Text & "'"
    Conexao.Execute (SQL)
    limp frmProdutos
    cmdIncluir.Enabled = False
    cmdExcluir.Enabled = False
    cmdAlterar.Enabled = False
    txtCod.Enabled = True
    fmDad.Enabled = False
    SQL = "Select * from Setor"
    Set tabeladinamica = Conexao.Execute(SQL)
    If Not tabeladinamica.EOF Then
        cmbSeto.Clear
        Do While Not tabeladinamica.EOF
            cmbSeto.AddItem tabeladinamica("Setor")
            tabeladinamica.MoveNext
        Loop
    End If
    SQL = "Select * from Fornecedores"
    Set tabeladinamica = Conexao.Execute(SQL)
    If Not tabeladinamica.EOF Then
        cmbForn.Clear
        Do While Not tabeladinamica.EOF
            cmbForn.AddItem tabeladinamica("Nome")
            tabeladinamica.MoveNext
        Loop
    End If
    SQL = "Select * from Produtos"
    Set tabeladinamica = Conexao.Execute(SQL)
    If Not tabeladinamica.EOF Then
        lstProd.Clear
        Do While Not tabeladinamica.EOF
            lstProd.AddItem tabeladinamica("Código") & " - " & tabeladinamica("Descricao")
            tabeladinamica.MoveNext
        Loop
    End If
End If
End Sub

Private Sub cmdExcluir_Click()
SQL = "Delete from Produtos where Código ='" & txtCod.Text & "'"
Conexao.Execute (SQL)
SQL = "Select * from Produtos"
Set tabeladinamica = Conexao.Execute(SQL)
If Not tabeladinamica.EOF Then
    lstProd.Clear
    Do While Not tabeladinamica.EOF
        lstProd.AddItem tabeladinamica("Código") & " - " & tabeladinamica("Descricao")
        tabeladinamica.MoveNext
    Loop
End If
SQL = "Select * from Setor"
Set tabeladinamica = Conexao.Execute(SQL)
If Not tabeladinamica.EOF Then
    cmbSeto.Clear
    Do While Not tabeladinamica.EOF
        cmbSeto.AddItem tabeladinamica("Setor")
        tabeladinamica.MoveNext
    Loop
End If
SQL = "Select * from Fornecedores"
Set tabeladinamica = Conexao.Execute(SQL)
If Not tabeladinamica.EOF Then
    cmbForn.Clear
    Do While Not tabeladinamica.EOF
        cmbForn.AddItem tabeladinamica("Nome")
        tabeladinamica.MoveNext
    Loop
End If
txtCod.Text = ""
txtDesc.Text = ""
txtQTDE.Text = ""
txtValo.Text = ""
cmdIncluir.Enabled = False
cmdExcluir.Enabled = False
cmdAlterar.Enabled = False
txtCod.Enabled = True
fmDad.Enabled = False
End Sub

Private Sub cmdIncluir_Click()
If txtDesc.Text = "" Then
    MsgBox ("Preencha a descrição")
ElseIf IsNumeric(txtDesc.Text) Then
    MsgBox ("Preencha a descrição corretamente"), vbExclamation, "Erro no cadastro"
ElseIf cmbSeto.Text = "" Then
    MsgBox ("Selecione o setor"), vbExclamation, "Erro no cadastro"
ElseIf txtQTDE.Text = "" Then
    MsgBox ("Preencha a quantidade em estoque"), vbExclamation, "Erro no cadastro"
ElseIf Not IsNumeric(txtQTDE.Text) Then
    MsgBox ("Insira apenas números na quantidade em estoque"), vbExclamation, "Erro no cadastro"
ElseIf txtValo.Text = "" Then
    MsgBox ("Insira o valor do produto"), vbExclamation, "Erro no cadastro"
ElseIf Not IsNumeric(txtValo.Text) Then
    MsgBox ("Insira apenas números no valor do produto"), vbExclamation, "Erro no cadastro"
ElseIf cmbForn.Text = "" Then
    MsgBox ("Selecione o fornecedor"), vbExclamation, "Erro no cadastro"
Else
    SQL = "Insert into Produtos(Código, Descricao, Setor, Fornecedor, QTDEstoque, Valor) Values('" & txtCod.Text & "','" & txtDesc.Text & "','" & cmbSeto.Text & "','" & cmbForn.Text & "','" & txtQTDE.Text & "','" & txtValo.Text & "')"
    Conexao.Execute (SQL)
    txtCod.Text = ""
    txtDesc.Text = ""
    txtQTDE.Text = ""
    txtValo.Text = ""
    cmdIncluir.Enabled = False
    cmdExcluir.Enabled = False
    cmdAlterar.Enabled = False
    txtCod.Enabled = True
    fmDad.Enabled = False
    SQL = "Select * from Setor"
    Set tabeladinamica = Conexao.Execute(SQL)
    If Not tabeladinamica.EOF Then
        cmbSeto.Clear
        Do While Not tabeladinamica.EOF
            cmbSeto.AddItem tabeladinamica("Setor")
            tabeladinamica.MoveNext
        Loop
    End If
    SQL = "Select * from Fornecedores"
    Set tabeladinamica = Conexao.Execute(SQL)
    If Not tabeladinamica.EOF Then
        cmbForn.Clear
        Do While Not tabeladinamica.EOF
            cmbForn.AddItem tabeladinamica("Nome")
            tabeladinamica.MoveNext
        Loop
    End If
    SQL = "Select * from Produtos"
    Set tabeladinamica = Conexao.Execute(SQL)
    If Not tabeladinamica.EOF Then
        lstProd.Clear
        Do While Not tabeladinamica.EOF
            lstProd.AddItem tabeladinamica("Código") & " - " & tabeladinamica("Descricao")
            tabeladinamica.MoveNext
        Loop
    End If
End If
End Sub

Private Sub cmdLimpar_Click()
cmdIncluir.Enabled = False
cmdExcluir.Enabled = False
cmdAlterar.Enabled = False
fmDad.Enabled = False
fmCod.Enabled = True
txtCod.Text = ""
lstProd.Clear
SQL = "Select * from Setor"
Set tabeladinamica = Conexao.Execute(SQL)
If Not tabeladinamica.EOF Then
    Do While Not tabeladinamica.EOF
        cmbSeto.AddItem tabeladinamica("Setor")
        tabeladinamica.MoveNext
    Loop
End If
SQL = "Select * from Fornecedores"
Set tabeladinamica = Conexao.Execute(SQL)
If Not tabeladinamica.EOF Then
    Do While Not tabeladinamica.EOF
        cmbForn.AddItem tabeladinamica("Nome")
        tabeladinamica.MoveNext
    Loop
End If
SQL = "Select * from Produtos"
Set tabeladinamica = Conexao.Execute(SQL)
If Not tabeladinamica.EOF Then
    Do While Not tabeladinamica.EOF
        lstProd.AddItem tabeladinamica("Código") & " - " & tabeladinamica("Descricao")
        tabeladinamica.MoveNext
    Loop
End If
End Sub


Private Sub Form_Load()
SQL = "Select * from Fornecedores"
Set tabeladinamica = Conexao.Execute(SQL)
If Not tabeladinamica.EOF Then
    cmbForn.Clear
    Do While Not tabeladinamica.EOF
        cmbForn.AddItem tabeladinamica("Nome")
        tabeladinamica.MoveNext
    Loop
End If
SQL = "Select * from Produtos"
Set tabeladinamica = Conexao.Execute(SQL)
If Not tabeladinamica.EOF Then
    lstProd.Clear
    Do While Not tabeladinamica.EOF
        lstProd.AddItem tabeladinamica("Código") & " - " & tabeladinamica("Descricao")
        tabeladinamica.MoveNext
    Loop
End If
SQL = "Select * from Setor"
    Set tabeladinamica = Conexao.Execute(SQL)
    If Not tabeladinamica.EOF Then
        cmbSeto.Clear
        Do While Not tabeladinamica.EOF
            cmbSeto.AddItem tabeladinamica("Setor")
            tabeladinamica.MoveNext
        Loop
    End If
End Sub

Private Sub lstProd_Click()
Dim codigo As String
Dim posicao As String
posicao = InStr(1, lstProd, "-")
codigo = Left(lstProd, posicao - 2)
SQL = "Select * from produtos where Código = '" & codigo & "'"
Set tabela = Conexao.Execute(SQL)
If tabela.EOF Then
    MsgBox ("nenhum produto cadastrado")
Else
    txtCod.Text = tabela("Código")
    txtDesc.Text = tabela("Descricao")
    cmbSeto = tabela("Setor")
    cmbForn.Text = tabela("Fornecedor")
    txtQTDE.Text = tabela("QTDEstoque")
    txtValo.Text = tabela("Valor")
    fmCod.Enabled = False
    fmDad.Enabled = True
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = True
    cmdExcluir.Enabled = True
End If
End Sub


Private Sub mnuVolt_Click()
Unload frmProdutos
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Len(txtCod.Text) <> 0 And IsNumeric(txtCod.Text) Then
    If VerCod("produtos", txtCod.Text, "Código") = False Then
        MsgBox ("Código já cadastrado")
    Else
        fmCod.Enabled = False
        fmDad.Enabled = True
        cmdIncluir.Enabled = True
    End If
End If
End Sub
