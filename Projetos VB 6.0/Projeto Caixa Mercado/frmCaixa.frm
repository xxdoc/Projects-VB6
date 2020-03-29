VERSION 5.00
Begin VB.Form frmCaixa 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Terminal Caixa"
   ClientHeight    =   9345
   ClientLeft      =   10875
   ClientTop       =   4785
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleMode       =   0  'User
   ScaleWidth      =   7800
   Begin VB.ListBox lstItens 
      Height          =   2205
      Left            =   5520
      TabIndex        =   19
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Frame fmProd 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   1980
      Left            =   480
      TabIndex        =   9
      Top             =   3975
      Width           =   6390
      Begin VB.Label lblSetor 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   2670
         TabIndex        =   15
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblValor 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00808000&
         Height          =   180
         Left            =   4260
         TabIndex        =   14
         Top             =   180
         Width           =   1440
      End
      Begin VB.Label lblNome 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00808000&
         Height          =   240
         Left            =   915
         TabIndex        =   13
         Top             =   195
         Width           =   1185
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Setor:"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   1965
         TabIndex        =   12
         Top             =   1050
         Width           =   795
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:  R$"
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   3540
         TabIndex        =   11
         Top             =   165
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         ForeColor       =   &H00808000&
         Height          =   225
         Left            =   225
         TabIndex        =   10
         Top             =   165
         Width           =   1050
      End
   End
   Begin VB.Frame fmCod 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Caption         =   "Código"
      Height          =   1245
      Left            =   270
      TabIndex        =   4
      Top             =   1455
      Width           =   3555
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   390
         TabIndex        =   5
         Top             =   600
         Width           =   2205
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         ForeColor       =   &H00808000&
         Height          =   225
         Left            =   420
         TabIndex        =   6
         Top             =   270
         Width           =   1560
      End
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00808000&
      Height          =   330
      Left            =   6870
      TabIndex        =   18
      Top             =   6840
      Width           =   2265
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Total da Compra:  R$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   330
      Left            =   4590
      TabIndex        =   17
      Top             =   6810
      Width           =   2325
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 - Quantidade     F3 - Fechar Pedido Enter - Adicionar produto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1470
      Left            =   360
      TabIndex        =   16
      Top             =   7200
      Width           =   2565
   End
   Begin VB.Label lblQTD 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00808000&
      Height          =   210
      Left            =   1965
      TabIndex        =   8
      Top             =   3345
      Width           =   840
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade"
      ForeColor       =   &H00808000&
      Height          =   270
      Left            =   765
      TabIndex        =   7
      Top             =   3330
      Width           =   825
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Caixa"
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
      Left            =   3855
      TabIndex        =   3
      Top             =   615
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Caixa"
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
      Left            =   3885
      TabIndex        =   2
      Top             =   615
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Terminal"
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
      Left            =   2565
      TabIndex        =   1
      Top             =   165
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Terminal"
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
      Left            =   2595
      TabIndex        =   0
      Top             =   165
      Width           =   2895
   End
   Begin VB.Menu mnuVolt 
      Caption         =   "Voltar"
   End
End
Attribute VB_Name = "frmCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub mnuVolt_Click()
Unload frmCaixa
End Sub



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
        Dim QLQC As String
        QLQC = InputBox("Digite a quantidade", "Terminal Caixa")
        If IsNumeric(QLQC) = True Then
            lblQTD.Caption = QLQC
        Else
            MsgBox ("digite uma quantidade númerica")
        End If
End If
If KeyCode = 13 And VerCod("produtos", txtCodigo, "Código") = False Then
    SQL = "Select * from Produtos Where Código = '" & txtCodigo & "'"
    Set tabela = Conexao.Execute(SQL)
    lblNome.Caption = tabela("Descricao")
    lblSetor.Caption = tabela("Setor")
    lblValor.Caption = tabela("Valor") * lblQTD.Caption
    lblTotal.Caption = Int(lblValor.Caption) + Int(lblTotal.Caption)
    txtCodigo.Text = ""
    lstItens.AddItem lblNome.Caption & " - " & lblQTD.Caption
End If
If KeyCode = 114 Then
    lstItens.Clear
    txtCodigo.Text = ""
    lblQTD.Caption = 1
    lblNome.Caption = ""
    lblValor.Caption = 0
    lblSetor.Caption = ""
    lblTotal.Caption = 0
End If
End Sub
