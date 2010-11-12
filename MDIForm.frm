VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   AutoShowChildren=   0   'False
   BackColor       =   &H00E0E0E0&
   Caption         =   "Arbimy manager 2.0 - Tricon Supermercados LTDA."
   ClientHeight    =   5655
   ClientLeft      =   4695
   ClientTop       =   1950
   ClientWidth     =   11400
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   5160
      Top             =   2640
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":339B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":36B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":17E23
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":29454
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":2976E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":2A048
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":700F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":82BDD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3555
      ScaleWidth      =   11340
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      Begin VB.Image Image1 
         Height          =   9240
         Left            =   2280
         Picture         =   "MDIForm.frx":864CB
         Top             =   -120
         Width           =   14355
      End
   End
   Begin VB.Menu mnu_cadpri 
      Caption         =   "Cadastros primários"
      Begin VB.Menu mnu_marcas 
         Caption         =   "Marcas"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnu_segmentos 
         Caption         =   "Segmentos"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_funcoes 
         Caption         =   "Funções"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnu_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_frete 
         Caption         =   "Preço do frete"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnu_1 
         Caption         =   "Formas de pagamentos"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnu_cadastros 
      Caption         =   "Cadastros"
      Begin VB.Menu mnu_clientes 
         Caption         =   "Clientes"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnu_fornecedores 
         Caption         =   "Fornecedores"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_funcionarios 
         Caption         =   "Funcionários"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnu_usuarios 
         Caption         =   "Usuários"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnu 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_contas 
         Caption         =   "Contas correntes"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_loca 
         Caption         =   "Localizações"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnu_produtos 
         Caption         =   "Produtos"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnu_estoque 
      Caption         =   "Controle de estoque"
   End
   Begin VB.Menu mnu_pedidos 
      Caption         =   "Pedidos"
      Begin VB.Menu mnu_compra 
         Caption         =   "Pedido de compra"
      End
   End
   Begin VB.Menu mnucalc 
      Caption         =   "Calculos"
   End
   Begin VB.Menu mnufdc 
      Caption         =   "Frente de Caixa"
   End
   Begin VB.Menu mnusai 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub MDIForm_Load()
            Call abrir_banco
            If Left(Time$, 2) > 18 And Left(Time$, 2) < 24 Then mensagem = " - BOA NOITE"
            If Left(Time$, 2) >= 0 And Left(Time$, 2) < 12 Then mensagem = " - BOM DIA"
            If Left(Time$, 2) > 12 And Left(Time$, 2) < 18 Then mensagem = " - BOA TARDE"
            
            MDIForm1.Caption = "TRICON SUPERMERCADOS LTDA." & "  |  " & Date$ & " - " & Time$ & mensagem
End Sub
Private Sub MDIForm_Resize()
            Picture1.Cls
            MDIForm1.Picture = LoadPicture("")
            Picture1.Visible = True
            Picture1.AutoRedraw = True
            Picture1.BackColor = &H8000000C
            Picture1.Height = Me.Height
            
            'Para centralizar a imagem no fundo
            
            Image1.Top = Picture1.Height / 2 - Image1.Height / 2
            Image1.Left = Picture1.Width / 2 - Image1.Width / 2
            
            'ou expandir a imagem por todo o fundo
            Image1.Stretch = True
            Image1.Top = 0
            Image1.Left = 0
            Image1.Height = Picture1.Height
            Image1.Width = Picture1.Width
            Picture1.PaintPicture Image1, Image1.Left, Image1.Top, Image1.Width, Image1.Height
            MDIForm1.Picture = Picture1.Image
            Picture1.Visible = False
End Sub

Private Sub mnu_separar_Click()

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
            End
End Sub

Private Sub mnu_1_Click()
            frm_formaspagamento.Show
End Sub

Private Sub mnu_clientes_Click()
            frm_clientes.Show
End Sub

Private Sub mnu_compra_Click()
            frm_pedidocompra.Show
End Sub

Private Sub mnu_contas_Click()
            frm_contas.Show
End Sub

Private Sub mnu_estoque_Click()
            frm_estoque.Show
End Sub

Private Sub mnu_fornecedores_Click()
            frm_fornecedores.Show
End Sub

Private Sub mnu_frete_Click()
            frm_frete.Show
End Sub

Private Sub mnu_funcionarios_Click()
            frm_funcionarios.Show
End Sub

Private Sub mnu_funcoes_Click()
            frm_funcao.Show
End Sub

Private Sub mnu_loca_Click()
            frm_localizacoes.Show
End Sub

Private Sub mnu_marcas_Click()
            frm_marcas.Show
End Sub

Private Sub mnu_produtos_Click()
            frm_produtos.Show
End Sub

Private Sub mnu_segmentos_Click()
            frm_segmentos.Show
End Sub

Private Sub mnu_usuarios_Click()
            frm_usuarios.Show
End Sub

Private Sub mnucalc_Click()
            form_estatistica.Show
End Sub

Private Sub mnufdc_Click()
            frm_frentedecaixa.Show
End Sub

Private Sub mnusai_Click()
            End
End Sub

Private Sub Timer1_Timer()
            MDIForm1.Caption = "Arbimy manager 2.0 - Tricon Supermercados LTDA.                               Bem vindo(a) Sr(a) " & usuario & ""
End Sub
