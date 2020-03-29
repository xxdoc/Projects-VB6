VERSION 5.00
Begin VB.MDIForm mdiMenu 
   BackColor       =   &H00FF80FF&
   Caption         =   "Menu Principal"
   ClientHeight    =   3705
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   5340
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuForm 
      Caption         =   "Formulários"
      Begin VB.Menu mnuProd 
         Caption         =   "Produtos"
      End
      Begin VB.Menu mnuFornecedores 
         Caption         =   "Fornecedores"
      End
      Begin VB.Menu fmSetores 
         Caption         =   "Setores"
      End
      Begin VB.Menu mnuClientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu mnuTC 
         Caption         =   "Terminal Caixa"
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "mdiMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub fmSetores_Click()
frmSetor.Show
End Sub

Private Sub mnuClientes_Click()
frmClientes.Show
End Sub

Private Sub mnuFornecedores_Click()
frmFornecedores.Show
End Sub

Private Sub mnuProd_Click()
frmProdutos.Show
End Sub

Private Sub mnuSair_Click()
End
End Sub

Private Sub mnuTC_Click()
frmCaixa.Show
End Sub
