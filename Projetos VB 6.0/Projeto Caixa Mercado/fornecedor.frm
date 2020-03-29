VERSION 5.00
Begin VB.Form frmfornecedor 
   BackColor       =   &H00FFC0FF&
   Caption         =   "FORNECEDORES"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   8610
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdlimpar 
      Caption         =   "LIMPAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   12
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton cmdincluir 
      Caption         =   "INCLUIR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   11
      Top             =   5790
      Width           =   1815
   End
   Begin VB.CommandButton cmdalterar 
      Caption         =   "ALTERAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   10
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton cmdexcluir 
      Caption         =   "EXCLUIR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   9
      Top             =   5760
      Width           =   2055
   End
   Begin VB.ListBox lstfornecedores 
      Height          =   1425
      Left            =   570
      TabIndex        =   8
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Frame fmdados 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Dados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   7335
      Begin VB.ComboBox lstestado 
         Height          =   288
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2160
         Width           =   1212
      End
      Begin VB.TextBox txtnome 
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtcidade 
         Height          =   375
         Left            =   5400
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtemail 
         Height          =   375
         Left            =   5400
         TabIndex        =   5
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtcontato 
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtendereco 
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "     E-MAIL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   22
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "  CONTATO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   21
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "  CIDADE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "ENDEREÇO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "   ESTADO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "       NOME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fmcnpj 
      BackColor       =   &H00FFC0FF&
      Caption         =   "CNPJ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   270
      TabIndex        =   0
      Top             =   690
      Width           =   1725
      Begin VB.TextBox txtcnpj 
         Height          =   495
         Left            =   300
         TabIndex        =   1
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.Label Label11 
      Caption         =   "   Fornecedores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   600
      TabIndex        =   24
      Top             =   5280
      Width           =   1305
   End
   Begin VB.Label Label7 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Label1"
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Menu Voltar 
      Caption         =   "Voltar"
   End
End
Attribute VB_Name = "frmfornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdalterar_Click()
If txtnome.Text = "" Or txtendereco.Text = "" Or txtcidade.Text = "" Or lstestado.Text = "" Or txtemail.Text = "" Or txtcontato.Text = "" Or IsNumeric(txtcontato.Text) = False Or IsNumeric(txtnome.Text) = True Or IsNumeric(txtcidade.Text) = True Then
    MsgBox "Preencha os campos corretamente", vbCritical, "Aviso do sistema"
Else
    SQL = "update fornecedores SET Nome = '" & txtnome.Text & "', Endereco = '" & txtendereco.Text & "', Cidade = '" & txtcidade.Text & "', Email = '" & txtemail.Text & "', Contato = '" & txtcontato.Text & "', Estado = '" & lstestado.Text & "' where CNPJ = '" & txtcnpj.Text & "'"
    CONEXAO.Execute (SQL)
    txtcnpj.Text = ""
    txtcidade.Text = ""
    txtendereco.Text = ""
    txtnome.Text = ""
    txtemail.Text = ""
    txtcontato.Text = ""
    fmcnpj.Enabled = True
    fmdados.Enabled = False
    cmdincluir.Enabled = False
    cmdexcluir.Enabled = False
    cmdalterar.Enabled = False
    cmdlimpar.Enabled = False
     SQL = "Select * from Estados"
    Set pegaestado = CONEXAO.Execute(SQL)
    If Not pegaestado.EOF Then
        lstestado.Clear
        Do While Not pegaestado.EOF
            lstestado.AddItem pegaestado("UF")
            pegaestado.MoveNext
        Loop
End If
SQL = "select * from fornecedores"
Set tabeladinamica = CONEXAO.Execute(SQL)
    If tabeladinamica.EOF Then
        MsgBox "Não existe nada cadastrado"
    Else
        lstfornecedores.Clear
            Do While Not tabeladinamica.EOF
                lstfornecedores.AddItem tabeladinamica("CNPJ") & "-" & tabeladinamica("Nome")
                    tabeladinamica.MoveNext
            Loop
    End If
End If
End Sub

Private Sub cmdexluir_Click()
If txtnome.Text = "" Or txtendereco.Text = "" Or txtcidade.Text = "" Or lstestado.Text = "" Or txtemail.Text = "" Or txtcontato.Text = "" Then
    MsgBox ("Existem campos em branco")
Else
    SQL = "delete from clientes where Nome = '" & txtnome.Text & "', Endereco = '" & txtendereco.Text & "', Cidade = '" & txtcidade.Text & "', Estado = '" & lstestado.Text & "', CNPJ = '" & txtcnpj.Text & "', Contato = '" & txtcontato.Text & ", Email = '" & txtemail.Text & "'"
    CONEXAO.Execute (SQL)
End If
End Sub

Private Sub cmdexcluir_Click()
If txtnome.Text = "" Or txtendereco.Text = "" Or txtcidade.Text = "" Or lstestado.Text = "" Then
    MsgBox ("Existem campos em branco")
Else
    SQL = "delete from fornecedores where CNPJ = '" & txtcnpj.Text & "'"
    CONEXAO.Execute (SQL)
    txtcnpj.Text = ""
    txtcidade.Text = ""
    txtendereco.Text = ""
    txtemail.Text = ""
    txtcontato.Text = ""
    txtnome.Text = ""
    fmcnpj.Enabled = True
    fmdados.Enabled = False
    cmdincluir.Enabled = False
    cmdexcluir.Enabled = False
    cmdalterar.Enabled = False
    cmdlimpar.Enabled = False
     SQL = "Select * from Estados"
    Set pegaestado = CONEXAO.Execute(SQL)
    If Not pegaestado.EOF Then
        lstestado.Clear
        Do While Not pegaestado.EOF
            lstestado.AddItem pegaestado("UF")
            pegaestado.MoveNext
        Loop
End If
SQL = "select * from fornecedores"
Set tabeladinamica = CONEXAO.Execute(SQL)
    If tabeladinamica.EOF Then
        MsgBox "Não existe nada cadastrado"
    Else
        lstfornecedores.Clear
            Do While Not tabeladinamica.EOF
                lstfornecedores.AddItem tabeladinamica("CNPJ") & "-" & tabeladinamica("Nome")
                    tabeladinamica.MoveNext
            Loop
    End If
End If
End Sub

Private Sub cmdincluir_Click()
If txtnome.Text = "" Or txtendereco.Text = "" Or txtcidade.Text = "" Or lstestado.Text = "" Or txtcontato = "" Or txtemail = "" Or IsNumeric(txtcidade.Text) = True Or IsNumeric(txtnome.Text) = True Or IsNumeric(txtcontato.Text) = False Then
    MsgBox "Preencha os campos corretamente", vbCritical, "Aviso do sistema"
Else
    SQL = "Insert into fornecedores (CNPJ, Nome, Endereco, Cidade, Contato, Email, Estado) Values ('" & txtcnpj.Text & "','" & txtnome.Text & "','" & txtendereco.Text & "','" & txtcidade.Text & "','" & txtcontato.Text & "','" & txtemail.Text & "','" & lstestado.Text & "')"
    CONEXAO.Execute (SQL)
    txtcnpj.Text = ""
    txtcidade.Text = ""
    txtendereco.Text = ""
    txtnome.Text = ""
    txtemail.Text = ""
    txtcontato.Text = ""
    fmcnpj.Enabled = True
    fmdados.Enabled = False
    cmdincluir.Enabled = False
    cmdexcluir.Enabled = False
    cmdalterar.Enabled = False
    cmdlimpar.Enabled = False
    lstestado.Clear
    lstestado.AddItem "SP"
lstestado.AddItem "RJ"
lstestado.AddItem "MG"
lstestado.AddItem "ES"
lstestado.AddItem "RS"
lstestado.AddItem "SC"
lstestado.AddItem "PR"
lstestado.AddItem "DF"
lstestado.AddItem "GO"
lstestado.AddItem "MS"
lstestado.AddItem "MT"
lstestado.AddItem "BA"
lstestado.AddItem "PE"
lstestado.AddItem "PB"
lstestado.AddItem "CE"
lstestado.AddItem "AL"
lstestado.AddItem "MA"
lstestado.AddItem "RN"
lstestado.AddItem "SE"
lstestado.AddItem "PI"
lstestado.AddItem "RR"
lstestado.AddItem "RO"
lstestado.AddItem "AM"
lstestado.AddItem "PA"
lstestado.AddItem "AP"
lstestado.AddItem "AC"
lstestado.AddItem "TO"
End If
SQL = "select * from fornecedores"
Set tabeladinamica = CONEXAO.Execute(SQL)
    If tabeladinamica.EOF Then
        MsgBox "Não existe nada cadastrado"
    Else
        lstfornecedores.Clear
            Do While Not tabeladinamica.EOF
                lstfornecedores.AddItem tabeladinamica("CNPJ") & "-" & tabeladinamica("Nome")
                    tabeladinamica.MoveNext
            Loop
    End If
End Sub

Private Sub cmdlimpar_Click()
txtcnpj.Text = ""
txtcidade.Text = ""
txtendereco.Text = ""
txtnome.Text = ""
txtcontato.Text = ""
txtemail.Text = ""
fmcnpj.Enabled = True
fmdados.Enabled = False
cmdincluir.Enabled = False
cmdexcluir.Enabled = False
cmdalterar.Enabled = False
cmdlimpar.Enabled = False
lstestado.AddItem "SP"
lstestado.AddItem "RJ"
lstestado.AddItem "MG"
lstestado.AddItem "ES"
lstestado.AddItem "RS"
lstestado.AddItem "SC"
lstestado.AddItem "PR"
lstestado.AddItem "DF"
lstestado.AddItem "GO"
lstestado.AddItem "MS"
lstestado.AddItem "MT"
lstestado.AddItem "BA"
lstestado.AddItem "PE"
lstestado.AddItem "PB"
lstestado.AddItem "CE"
lstestado.AddItem "AL"
lstestado.AddItem "MA"
lstestado.AddItem "RN"
lstestado.AddItem "SE"
lstestado.AddItem "PI"
lstestado.AddItem "RR"
lstestado.AddItem "RO"
lstestado.AddItem "AM"
lstestado.AddItem "PA"
lstestado.AddItem "AP"
lstestado.AddItem "AC"
lstestado.AddItem "TO"
End Sub

Private Sub fmdados_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Form_Load()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub lstfornecedores_Click()

posicao = InStr(1, lstfornecedores, "-")
cnpj = Left(lstfornecedores, posicao - 1)
    SQL = "select * from Fornecedores where CNPJ = '" & (cnpj) & "'"
        Set tabela = CONEXAO.Execute(SQL)
            If tabela.EOF Then
                MsgBox ("Nada cadastrado")
            Else
                txtcnpj.Text = tabela("CNPJ")
                txtnome.Text = tabela("Nome")
                txtendereco.Text = tabela("Endereco")
                txtcidade.Text = tabela("Cidade")
                txtemail.Text = tabela("Email")
                txtcontato.Text = tabela("Contato")
                lstestado.Text = tabela("Estado")
                fmcnpj.Enabled = False
                fmdados.Enabled = True
                cmdincluir.Enabled = False
                cmdalterar.Enabled = True
                cmdexcluir.Enabled = True
            End If
End Sub


Private Sub txtcnpj_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And IsNumeric(txtcnpj.Text) = True Then
    SQL = "Select * from fornecedores where CNPJ ='" & txtcnpj.Text & "'"
    Set tabeladinamica = CONEXAO.Execute(SQL)
        If tabeladinamica.EOF Then
            fmcnpj.Enabled = False
            fmdados.Enabled = True
            cmdincluir.Enabled = True
            cmdalterar.Enabled = False
            cmdexcluir.Enabled = False
            cmdlimpar.Enabled = True
        End If
End If
End Sub

Private Sub Voltar_Click()
Unload frmfornecedor
End Sub
