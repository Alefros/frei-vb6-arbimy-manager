VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_produtos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de produtos"
   ClientHeight    =   7635
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cbo_fornecedor 
      Height          =   315
      Left            =   6120
      TabIndex        =   53
      Top             =   600
      Width           =   3135
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Preço"
      Height          =   1335
      Left            =   0
      TabIndex        =   44
      Top             =   3960
      Width           =   5535
      Begin VB.TextBox txt_icms 
         Height          =   285
         Left            =   4200
         TabIndex        =   52
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txt_lucro 
         Height          =   285
         Left            =   2160
         TabIndex        =   50
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txt_custo 
         Height          =   285
         Left            =   2160
         TabIndex        =   48
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox txt_valor_u 
         Height          =   285
         Left            =   2160
         TabIndex        =   46
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "ICMS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   51
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Margem de lucro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Preço de custo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Valor unitário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informações dos fornecedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4560
      TabIndex        =   28
      Top             =   360
      Width           =   4815
      Begin VB.TextBox txt_razaosocial 
         Height          =   285
         Left            =   1560
         TabIndex        =   29
         Top             =   600
         Width           =   3135
      End
      Begin MSMask.MaskEdBox msk_ie 
         Height          =   255
         Left            =   1560
         TabIndex        =   30
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   15
         Mask            =   "999.999.999.999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_cnpj 
         Height          =   255
         Left            =   1560
         TabIndex        =   31
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   18
         Mask            =   "99.999.999/9999-99"
         PromptChar      =   "_"
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Razão social"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "IE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CNPJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Width           =   615
      End
   End
   Begin VB.TextBox txt_codigo_b 
      Height          =   285
      Left            =   6120
      MaxLength       =   14
      TabIndex        =   20
      Top             =   0
      Width           =   3255
   End
   Begin VB.TextBox txt_codigo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   18
      Top             =   0
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8760
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Comandos"
      Height          =   1335
      Left            =   5640
      TabIndex        =   16
      Top             =   3960
      Width           =   3735
      Begin VB.CommandButton cmd_excluir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Excluir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2760
         Picture         =   "frm_produtos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmd_alterar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Alterar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1800
         Picture         =   "frm_produtos.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Alterar"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmd_salvar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Salvar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   960
         Picture         =   "frm_produtos.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salvar"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmd_novo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Novo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Picture         =   "frm_produtos.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Novo"
         Top             =   360
         Width           =   735
      End
   End
   Begin MSFlexGridLib.MSFlexGrid mfg_produtos 
      Height          =   2295
      Left            =   0
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5280
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4048
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imagem"
      Height          =   1815
      Left            =   4560
      TabIndex        =   14
      Top             =   2160
      Width           =   4815
      Begin VB.CommandButton cmd_limpar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpar"
         DragIcon        =   "frm_produtos.frx":0C28
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         Picture         =   "frm_produtos.frx":106A
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   960
         Width           =   795
      End
      Begin VB.CommandButton cmd_procurar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Procurar..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.Image img_imagem 
         Height          =   1455
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1815
      Left            =   0
      TabIndex        =   11
      Top             =   2160
      Width           =   4455
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Especificações"
         Height          =   495
         Left            =   0
         TabIndex        =   40
         Top             =   1320
         Width           =   4455
         Begin VB.OptionButton opt_pacote 
            BackColor       =   &H00E0E0E0&
            Caption         =   "PACOTE"
            Height          =   195
            Left            =   2040
            TabIndex        =   42
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton opt_caixa 
            BackColor       =   &H00E0E0E0&
            Caption         =   "CAIXA"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.OptionButton opt_perecivel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Perecível"
         Height          =   495
         Left            =   1320
         TabIndex        =   39
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Não perecível"
         Height          =   495
         Left            =   3120
         TabIndex        =   38
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txt_peso 
         Height          =   285
         Left            =   1320
         TabIndex        =   37
         Top             =   120
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtp_validade 
         Height          =   255
         Left            =   1320
         TabIndex        =   17
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         DateIsNull      =   -1  'True
         Format          =   94633985
         CurrentDate     =   40489
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Unidade"
         Height          =   495
         Left            =   2760
         TabIndex        =   12
         Top             =   0
         Width           =   1695
         Begin VB.OptionButton opt_grama 
            BackColor       =   &H00E0E0E0&
            Caption         =   "G"
            Height          =   195
            Left            =   1200
            TabIndex        =   4
            Top             =   240
            Width           =   375
         End
         Begin VB.OptionButton opt_litro 
            BackColor       =   &H00E0E0E0&
            Caption         =   "L"
            Height          =   195
            Left            =   720
            TabIndex        =   3
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton opt_quilo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Kg"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Peso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Validade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4455
      Begin VB.ComboBox cbo_segmento 
         Height          =   315
         Left            =   1320
         TabIndex        =   27
         Top             =   1320
         Width           =   3015
      End
      Begin VB.ComboBox cbo_marca 
         Height          =   315
         Left            =   1320
         TabIndex        =   24
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   23
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Segmento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Marca"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descrição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cód de barras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   21
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "frm_produtos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim codprodutos As Integer
Public tabsegmentos As New ADODB.Recordset
Public tabprodutos As New ADODB.Recordset
Dim L_linha, l_codprodutos, l_marca, l_fornecedor, l_segmento, veri As Integer
Dim unidade, especificacao As String

Private Sub cbo_fornecedor_Click()
            l_fornecedor = cbo_fornecedor.ItemData(cbo_fornecedor.ListIndex)
End Sub

Private Sub cbo_marca_Click()
            
            'Call carregar_combo
            
            l_marca = cbo_marca.ItemData(cbo_marca.ListIndex)
            
End Sub

Private Sub cbo_segmento_Click()
            'Call carregar_combo
            l_segmento = cbo_segmento.ItemData(cbo_segmento.ListIndex)
End Sub
Private Sub cmd_alterar_Click()
            status = "alteradas"
            veri = 0
            Call verificar
            If veri <> 1 Then
            veri = 0
            Call gravar_produtos
            If veri <> 1 Then Call box
            Call flex
            End If
End Sub
Private Sub cmd_excluir_Click()
                status = "Excluídas"
        If MsgBox("DESEJA REALMENTE EXCLUIR ESTE PRODUTO ?", vbYesNo + vbDefaultButton2 + vbQuestion, "TRICON SUPERMERCADOS LTDA.") = vbYes Then
                If tabprodutos.State = adStateOpen Then tabprodutos.Close
                tabprodutos.Open "Select * from Produtos where Codigo =" & txt_codigo
                    If tabprodutos.RecordCount <> 0 Then
                    conectar.Execute "delete from Produtos where Codigo =" & txt_codigo
                    Call flex
                    Call box
                    Call cod_produtos
                    cmd_novo_Click
                    Else
                    MsgBox "ESTE PRODUTO NÃO EXISTE", vbInformation, "TRICON SUPERMERCADOS LTDA."
                    txt_banco.SetFocus
                    End If
        End If
        Exit Sub
End Sub

Private Sub cmd_limpar_Click()
            img_imagem.Picture = LoadPicture(Empty)
End Sub

Private Sub cmd_novo_Click()
            Call limpar
            Call cod_produtos
End Sub
Private Sub cmd_procurar_Click()
            CommonDialog1.DialogTitle = "PROCURAR IMAGEM - TRICON SUPERMERCADOS LTDA."
            CommonDialog1.Filter = "Bitmap (*bmp; *.dib)|*bmp; *.dib|JPEG (*.jpg; *.jpeg; *.jpe; *.jfif)|*.jpg; *.jpeg; *.jpe; *.jfif|GIF (*.gif)|*.gif|PNG (*.png)|*.png|TIFF (*.tif; *.tiff)|*.tif; *.tiff|Todas as Imagens|*bmp; *.dib; *.jpg; *.jpeg; *.jpe; *.jfif; *.gif; *.png; *.tif; *.tiff;"
            CommonDialog1.FilterIndex = 6
            CommonDialog1.ShowOpen
            If CommonDialog1.FileName <> "" Then
            img_imagem.Picture = LoadPicture(CommonDialog1.FileName)
            End If
            
            
            
            'CommonDialog1.DialogTitle = "PROCURAR IMAGEM - TRICON SUPERMERCADOS LTDA."
            'CommonDialog1.Filter = "Bitmap (*bmp; *.dib)|*bmp; *.dib|JPEG (*.jpg; *.jpeg; *.jpe; *.jfif)|*.jpg; *.jpeg; *.jpe; *.jfif|GIF (*.gif)|*.gif|PNG (*.png)|*.png|TIFF (*.tif; *.tiff)|*.tif; *.tiff|Todas as Imagens|*bmp; *.dib; *.jpg; *.jpeg; *.jpe; *.jfif; *.gif; *.png; *.tif; *.tiff;"
            'CommonDialog1.ShowOpen
            'If CommonDialog1.FileName <> "" Then
            'img_imagem.Picture = LoadPicture(CommonDialog1.FileName)
            'End If
End Sub

Private Sub carregar_combo()
            If tabmarca.State = adStateOpen Then tabmarca.Close
            tabmarca.Open "Marcas", conectar, adOpenKeyset, adLockOptimistic
            
            If tabfornecedores.State = adStateOpen Then tabfornecedores.Close
            tabfornecedores.Open "Fornecedores", conectar, adOpenKeyset, adLockOptimistic
            
            If tabsegmentos.State = adStateOpen Then tabsegmentos.Close
            tabsegmentos.Open "Segmentos", conectar, adOpenKeyset, adLockOptimistic
            
            Do Until tabmarca.EOF = True
            
            cbo_marca.AddItem tabmarca!Marca
            cbo_marca.ItemData(cbo_marca.NewIndex) = tabmarca!codigo
            tabmarca.MoveNext
            Loop
            
            Do Until tabfornecedores.EOF = True
            cbo_fornecedor.AddItem tabfornecedores!Razao_social
            cbo_fornecedor.ItemData(cbo_fornecedor.NewIndex) = tabfornecedores!codigo
            tabfornecedores.MoveNext
            Loop
            
            Do Until tabsegmentos.EOF = True
            cbo_segmento.AddItem tabsegmentos!Segmento
            cbo_segmento.ItemData(cbo_segmento.NewIndex) = tabsegmentos!codigo
            tabsegmentos.MoveNext
            Loop
End Sub

Private Sub cmd_salvar_Click()
            status = "gravadas"
            veri = 0
            Call verificar
    
            If veri <> 1 Then
            veri = 0
            Call gravar_produtos
            If veri <> 1 Then Call box
            If veri <> 1 Then Call flex
            End If
End Sub

Private Sub Form_Load()
            Call abrir_banco
            If tabmarca.State = adStateOpen Then tabmarca.Close
            tabmarca.Open "Marcas", conectar, adOpenKeyset, adLockOptimistic
            
            If tabfornecedores.State = adStateOpen Then tabfornecedores.Close
            tabfornecedores.Open "Fornecedores", conectar, adOpenKeyset, adLockOptimistic
            
            If tabsegmentos.State = adStateOpen Then tabsegmentos.Close
            tabsegmentos.Open "Segmentos", conectar, adOpenKeyset, adLockOptimistic
            
            If tabprodutos.State = adStateOpen Then tabprodutos.Close
            tabprodutos.Open "Produtos", conectar, adOpenKeyset, adLockOptimistic
            
            Call carregar_combo
            Call cod_produtos
            Call flex
End Sub
Private Sub cod_produtos()
            codprodutos = 1
a:
            If tabprodutos.State = adStateOpen Then tabprodutos.Close
            tabprodutos.Open "select * from Produtos where Codigo =" & codprodutos
            If tabprodutos.RecordCount > 0 Then
                codprodutos = codprodutos + 1
                GoTo a:
            End If
            txt_codigo = codprodutos
End Sub

Private Sub flex()
            If tabprodutos.State = adStateOpen Then tabprodutos.Close
            tabprodutos.Open "Produtos", conectar, adOpenKeyset, adLockOptimistic
            
            If tabprodutos!codigo <> "" Then tabprodutos.MoveFirst
            mfg_produtos.Clear
            mfg_produtos.Rows = 2
            mfg_produtos.FormatString = "Código    |  Nome                                                            |Validade          |Peso              |Unidade      |Valor Untário      "
            Do Until tabprodutos.EOF = True
            mfg_produtos.TextMatrix(mfg_produtos.Rows - 1, 0) = tabprodutos!codigo
            mfg_produtos.TextMatrix(mfg_produtos.Rows - 1, 1) = tabprodutos!nome
            mfg_produtos.TextMatrix(mfg_produtos.Rows - 1, 2) = tabprodutos!Validade
            mfg_produtos.TextMatrix(mfg_produtos.Rows - 1, 3) = tabprodutos!peso
            mfg_produtos.TextMatrix(mfg_produtos.Rows - 1, 4) = tabprodutos!unidade
            mfg_produtos.TextMatrix(mfg_produtos.Rows - 1, 5) = Format(tabprodutos!Preco_unitario, "currency")
            mfg_produtos.Rows = mfg_produtos.Rows + 1
            tabprodutos.MoveNext
            Loop
            If mfg_produtos.Rows <> 2 Then mfg_produtos.Rows = mfg_produtos.Rows - 1
End Sub
Private Sub gravar_produtos()

            If status <> "alteradas" Then
            Call abrir_banco
            If tabprodutos.State = adStateOpen Then tabprodutos.Close
            tabprodutos.Open "Produtos", conectar, adOpenKeyset, adLockOptimistic
            tabprodutos.AddNew
            Else
            If tabprodutos.State = adStateOpen Then tabprodutos.Close
            tabprodutos.Open "select * from Produtos where Codigo =" & txt_codigo
            If tabprodutos.RecordCount = 0 Then
            MsgBox "Este produto não esta cadastrado", vbInformation, "Arbimy manager 2.0"
            'veri = 1
            Exit Sub
            End If
            End If
            
            'If tabcodigobarra.State = adStateOpen Then tabcodigobarra.Close
            'tabcodigobarra.Open "Codigo_barra", conectar, adOpenKeyset, adLockOptimistic
            'tabcodigobarra.Close
            'tabcodigobarra.Open "select * from Codigo_barra where Codigo_barra = '" & txt_codigo_b & "'"
            'If tabcodigobarra.RecordCount = 0 Then
           ' MsgBox "Código de barras não existente", vbInformation, "Arbimy manager 2.0"
            'veri = 1
            'Exit Sub
            'Else
            'codigo2 = tabcodigobarra!codigo
           ' End If
            
            If opt_grama = True Then unidade = opt_grama.Caption
            If opt_litro = True Then unidade = opt_litro.Caption
            If opt_quilo = True Then unidade = opt_quilo.Caption
            If opt_caixa = True Then especificação = opt_caixa.Caption
            If opt_pacote = True Then especificação = opt_pacote.Caption
            If status <> "alteradas" Then tabprodutos!codigo = txt_codigo
            
            tabprodutos!Codigo_barras = txt_codigo_b
            tabprodutos!codigo = txt_codigo
            tabprodutos!nome = txt_nome
            tabprodutos!peso = txt_peso
            tabprodutos!unidade = unidade
            tabprodutos!Validade = dtp_validade
            
            tabprodutos!Imagem = CommonDialog1.FileName
           
            
            tabprodutos!Preco_unitario = txt_valor_u
            tabprodutos!especificacao = especificacao
            tabprodutos!Marca = l_marca
            tabprodutos!Fornecedor = l_fornecedor
            tabprodutos!Segmento = l_segmento
            tabprodutos.Update
End Sub
Private Sub verificar()
            If txt_codigo_b.Text = "" Then MsgBox "INSIRA O CÓDIGO DE BARRAS DO PRODUTO", vbInformation, "TRICON SUPERMERCADOS LTDA.": veri = 1: Exit Sub
            If txt_nome.Text = "" Then MsgBox "DIGITE O NOME DO PRODUTO", vbInformation, "TRICON SUPERMERCADOS LTDA.": veri = 1: Exit Sub
            If cbo_fornecedor.Text = "" Then MsgBox "SELECIONE O FORNECEDOR DO PRODUTO", vbInformation, "TRICON SUPERMERCADOS LTDA.": veri = 1: Exit Sub
            If cbo_marca.Text = "" Then MsgBox "SELECIONE A MARCA DO PRODUTO", vbInformation, "TRICON SUPERMERCADOS LTDA.": veri = 1: Exit Sub
            If cbo_segmento.Text = "" Then MsgBox "SELECIONE O SEGMENTO DO PRODUTO", vbInformation, "TRICON SUPERMERCADOS LTDA.": veri = 1: Exit Sub
            If txt_peso.Text = "" Then MsgBox "DIGITE O PESO DO PRODUTO", vbInformation, "TRICON SUPERMERCADOS LTDA.": veri = 1: Exit Sub
            If txt_valor_u.Text = "" Then MsgBox "DIGITE O VALOR UNITÁRIO DO PRODUTO", vbInformation, "TRICON SUPERMERCADOS LTDA.": veri = 1: Exit Sub
            If opt_grama = False And opt_litro = False And opt_quilo = False Then MsgBox "SELECIONE A UNIDADE DO PRODUTO", vbInformation, "TRICON SUPERMERCADOS LTDA.": veri = 1: Exit Sub
            If opt_caixa = False And opt_pacote = False Then MsgBox "SELECIONE UMA ESPECIFICAÇÃO DO PRODUTO", vbInformation, "TRICON SUPERMERCADOS LTDA.": veri = 1: Exit Sub
End Sub
Private Sub Form_Unload(Cancel As Integer)
            Call voltar_botao
End Sub

Private Sub mfg_produtos_Click()
            On Error GoTo ERRO:
            If mfg_produtos.Rows < 2 Then Exit Sub
            L_linha = mfg_produtos.Row
            l_codprodutos = mfg_produtos.TextMatrix(L_linha, 0)
            If tabprodutos.State = adStateOpen Then tabprodutos.Close
            tabprodutos.Open "select * from Produtos where Codigo =" & l_codprodutos
            If tabprodutos.RecordCount > 0 Then
            txt_codigo = tabprodutos!codigo
            txt_codigo_b = tabprodutos!Codigo_barras
            txt_nome = tabprodutos!nome
            txt_peso = tabprodutos!peso
            If tabprodutos!unidade = opt_grama.Caption Then opt_grama = True
            If tabprodutos!unidade = opt_litro.Caption Then opt_litro = True
            If tabprodutos!unidade = opt_quilo.Caption Then opt_quilo = True
            If tabprodutos!especificacao = opt_caixa.Caption Then opt_caixa = True
            If tabprodutos!especificacao = opt_pacote.Caption Then opt_pacote = True
            cbo_marca.ListIndex = tabprodutos!Marca - 1
            cbo_fornecedor.ListIndex = tabprodutos!Fornecedor - 1
            cbo_segmento.ListIndex = tabprodutos!Segmento - 1
            dtp_validade.value = tabprodutos!Validade
            CommonDialog1.FileName = tabprodutos!Imagem
            txt_valor_u = Format(tabprodutos!Preco_unitario, "currency")
            img_imagem.Picture = LoadPicture(tabprodutos!Imagem)
            End If
            Exit Sub
ERRO:
            txt_codigo_b = Err.Description
            If Err.Description = "File not found: 'C:\Users\User\Desktop\p_presunto  perdigao.bmp'" Then
            img_imagem.Picture = LoadPicture(App.Path & "\imagem erro.jpg")
            End If
End Sub

Private Sub txt_valor_u_LostFocus()
txt_valor_u = Format(txt_valor_u, "currency")
End Sub
Private Sub Desabilitar()
            msk_ie.PromptInclude = False
            msk_cnpj.PromptInclude = False
End Sub
Private Sub Habilitar()
            msk_ie.PromptInclude = True
            msk_cnpj.PromptInclude = True
End Sub
Private Sub limpar()
            Call Desabilitar
            txt_codigo_b = Clear
            txt_descricao = Clear
            msk_ie.Text = Clear
            txt_custo = Clear
            txt_lucro = Clear
            txt_icms = Clear
            msk_cnpj.Text = Clear
            txt_razaosocial = Clear
            txt_nome = Clear
            cbo_fornecedor = ""
            cbo_marca = ""
            cbo_segmento = ""
            txt_descricao = Clear
            txt_peso = Clear
            txt_valor_u = Clear
            opt_quilo = False
            opt_litro = False
            opt_grama = False
            opt_caixa = False
            opt_pacote = False
            img_imagem.Picture = LoadPicture(Empty)
            txt_codigo_b.SetFocus
            Call Habilitar
End Sub
