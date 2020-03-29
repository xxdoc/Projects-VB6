VERSION 5.00
Begin VB.Form frmAutenticar 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autenticação"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3960
   ControlBox      =   0   'False
   FillColor       =   &H00FFC0FF&
   ForeColor       =   &H00FFC0FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAutenticar 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Autenticar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1170
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   5
      Top             =   2265
      Width           =   975
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   2385
      MaskColor       =   &H8000000F&
      TabIndex        =   6
      Top             =   2265
      Width           =   975
   End
   Begin VB.TextBox txtSenha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1410
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1785
      Width           =   1575
   End
   Begin VB.TextBox txtLogin 
      Height          =   285
      Left            =   1410
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1425
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Autenticação"
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
      Left            =   420
      TabIndex        =   7
      Top             =   210
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   690
      TabIndex        =   3
      Top             =   1785
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Login:"
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   690
      TabIndex        =   1
      Top             =   1425
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Autenticação"
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
      Height          =   615
      Left            =   450
      TabIndex        =   0
      Top             =   225
      Width           =   3135
   End
End
Attribute VB_Name = "frmAutenticar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAutenticar_Click()
SQL = "Select * from autenticacao where login = '" & txtLogin.Text & "'and Senha= '" & txtSenha.Text & "'"
Set tabeladinamica = Conexao.Execute(SQL)
If tabeladinamica.EOF Then
    MsgBox ("Login ou senha não encontrado")
Else
    mdiMenu.Show
    Unload frmAutenticar
End If
End Sub

Private Sub cmdSair_Click()
End
End Sub
Private Sub Form_Load()
Conexao.Open "PROVIder = Microsoft.jet.oledb.4.0; data source = E:\Informatica\Informática I40 Semestre 2\LPR\Projeto\Banco\Dados.mdb"
End Sub

Private Sub txtLogin_Change()
If txtLogin.Text <> "" And txtSenha.Text <> "" Then
    cmdAutenticar.Enabled = True
Else
    cmdAutenticar.Enabled = False
End If
End Sub

Private Sub txtLogin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtLogin.Text <> "" And txtSenha.Text <> "" Then
    SQL = "Select * from autenticacao where login = '" & txtLogin.Text & "'and Senha= '" & txtSenha.Text & "'"
    Set tabeladinamica = Conexao.Execute(SQL)
    If tabeladinamica.EOF Then
        MsgBox ("Login ou senha não encontrado")
    Else
        mdiMenu.Show
        Unload frmAutenticar
    End If
End If
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And txtLogin.Text <> "" And txtSenha.Text <> "" Then
    SQL = "Select * from autenticacao where login = '" & txtLogin.Text & "' and Senha = '" & txtSenha.Text & "'"
    Set tabeladinamica = Conexao.Execute(SQL)
    If tabeladinamica.EOF Then
        MsgBox ("Login ou senha não encontrado")
    Else
        mdiMenu.Show
        Unload frmAutenticar
    End If
End If
End Sub

Private Sub txtSenha_Change()
If txtLogin.Text <> "" And txtSenha.Text <> "" Then
    cmdAutenticar.Enabled = True
Else
    cmdAutenticar.Enabled = False
End If
End Sub
