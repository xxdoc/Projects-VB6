VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Menu"
   ClientHeight    =   3180
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4800
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuForm 
      Caption         =   "Formulários"
      Begin VB.Menu mnuVendas 
         Caption         =   "Vendas"
      End
      Begin VB.Menu mnuS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTabSimp 
         Caption         =   "Tabuadas Simples"
         Begin VB.Menu mnuTabFor 
            Caption         =   "For"
         End
         Begin VB.Menu mnuTabWhile 
            Caption         =   "While"
         End
      End
      Begin VB.Menu mnuS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFtab 
         Caption         =   "Faixa tabuada"
         Begin VB.Menu mnuFtabFor 
            Caption         =   "For"
         End
         Begin VB.Menu mnuFtabWhile 
            Caption         =   "While"
         End
      End
      Begin VB.Menu mnuS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunc 
         Caption         =   "Funções"
         Begin VB.Menu mnuFuncLen 
            Caption         =   "Len"
         End
         Begin VB.Menu mnuFuncLeft 
            Caption         =   "Left"
         End
         Begin VB.Menu mnuFuncRight 
            Caption         =   "Right"
         End
         Begin VB.Menu mnuFuncMid 
            Caption         =   "Mid"
         End
         Begin VB.Menu mnuFuncInStr 
            Caption         =   "InStr"
         End
      End
      Begin VB.Menu mnuS4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuApp 
         Caption         =   "Aplicativos"
         Begin VB.Menu mnuAppNomSob 
            Caption         =   "Separar  nome e sobrenome"
         End
         Begin VB.Menu mnuAppSVog 
            Caption         =   "Separar vogais"
         End
         Begin VB.Menu mnuAppCVo 
            Caption         =   "Contar volgal"
         End
         Begin VB.Menu mnuAppDVo 
            Caption         =   "Contar dupla vogal"
         End
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAppCVo_Click()
Form13.Show
End Sub
Private Sub mnuAppNomSob_Click()
Form11.Show
End Sub
Private Sub mnuAppSVog_Click()
Form12.Show
End Sub
Private Sub mnuFtabFor_Click()
Form5.Show
End Sub
Private Sub mnuFtabWhile_Click()
Form4.Show
End Sub
Private Sub mnuFuncInStr_Click()
Form9.Show
End Sub
Private Sub mnuFuncLeft_Click()
Form7.Show
End Sub
Private Sub mnuFuncLen_Click()
Form6.Show
End Sub
Private Sub mnuFuncRight_Click()
Form8.Show
End Sub

Private Sub mnuSair_Click()
End
End Sub

Private Sub mnuTabFor_Click()
Form1.Show
End Sub
Private Sub mnuTabWhile_Click()
Form3.Show
End Sub
Private Sub mnuVendas_Click()
Form2.Show
End Sub
