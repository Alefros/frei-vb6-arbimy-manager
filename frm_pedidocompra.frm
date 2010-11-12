VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_pedidocompra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedido de Compra"
   ClientHeight    =   7530
   ClientLeft      =   -15
   ClientTop       =   360
   ClientWidth     =   9900
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_pedidocompra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Comandos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   4215
      Left            =   8760
      TabIndex        =   23
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Salvar"
         Enabled         =   0   'False
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
         Left            =   240
         Picture         =   "frm_pedidocompra.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Alterar"
         Enabled         =   0   'False
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
         Left            =   240
         Picture         =   "frm_pedidocompra.frx":0FD4
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Excluir"
         Enabled         =   0   'False
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
         Left            =   240
         Picture         =   "frm_pedidocompra.frx":12DE
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3240
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
         Left            =   240
         Picture         =   "frm_pedidocompra.frx":15E8
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Adicionar a lista..."
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
      Left            =   2760
      TabIndex        =   22
      Top             =   4200
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker dtp_data 
      Height          =   375
      Left            =   7200
      TabIndex        =   19
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   110362625
      CurrentDate     =   40493
   End
   Begin VB.TextBox txt_np 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   17
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Informações do fornecedor"
      Height          =   1815
      Left            =   5040
      TabIndex        =   11
      Top             =   2760
      Width           =   3615
      Begin VB.TextBox txt_rs 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txt_codfornecedor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.Label LABEL222 
         Alignment       =   1  'Right Justify
         Caption         =   "Razão Social"
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
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
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
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Informações do produto"
      Height          =   1815
      Left            =   0
      TabIndex        =   7
      Top             =   2760
      Width           =   4935
      Begin VB.TextBox txt_vu 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txt_descricao 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txt_total 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txt_qtde 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txt_codproduto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Unitário"
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
         Left            =   2760
         TabIndex        =   21
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Total"
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
         TabIndex        =   15
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantidade"
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
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de produtos"
      Height          =   2415
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   8655
      Begin MSFlexGridLib.MSFlexGrid mfg_produtos 
         Height          =   2055
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   3625
         _Version        =   393216
         FormatString    =   " "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Data"
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
      Left            =   5880
      TabIndex        =   18
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Numero do Pedido"
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
      TabIndex        =   16
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frm_pedidocompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l_cod_pedido_compra, var, L_linha As Integer

Private Sub cmd_alterar_Click()
            status = "alteradas"
            If txt_codfornecedor <> "" Or txt_codproduto <> "" Then
            var = 0
            Call gravar_pdc
            If var <> 1 Then Call box
            If var <> 1 Then Call flex
            End If
End Sub

Private Sub cmd_excluir_Click()
            status = "excluídas"
            If MsgBox("DESEJA REALMENTE EXCLUIR ESTE PEDIDO ?", vbYesNo + vbDefaultButton2 + vbQuestion, "TRICON SUPERMERCADOS LTDA.") = vbYes Then
                If tabpedidodecompra.State = adStateOpen Then tabpedidodecompra.Close
                tabpedidodecompra.Open "Select * from Pedidos where Codigo =" & txt_np
                    If tabsegmentos.RecordCount <> 0 Then
                    conectar.Execute "delete from Pedidos where Codigo =" & txt_np
                    Call flex
                    Call box
                    Call cod_pdc
                    cmd_novo_Click
                    Else
                    MsgBox "ESTE PEDIDO NÃO EXISTE", vbInformation, "TRICON SUPERMERCADOS LTDA."
                    txt_codproduto.SetFocus
                    End If
            End If
End Sub

Private Sub cmd_novo_Click()
Call desativar
txt_codproduto = Clear
txt_segmento = Clear
txt_nome = Clear
txt_marca = Clear
txt_unidade = Clear
txt_qtde = Clear
txt_vu = Clear
txt_total = Clear
txt_codfornecedor = Clear
txt_rs = Clear
msk_conta = Clear
msk_cnpj = Clear
msk_ie = Clear
txt_codproduto.SetFocus
Call ativar
Call cod_pdc
dtp_data.value = Date$
End Sub

Private Sub cmd_salvar_Click()
            status = "gravadas"
            If txt_codfornecedor <> "" Or txt_codproduto <> "" Then
            var = 0
            Call gravar_pdc
            If var <> 1 Then Call box
            If var <> 1 Then Call flex
            End If
End Sub

Private Sub Form_Load()

Call abrir_banco

If tabprodutos.State = adStateOpen Then tabprodutos.Close
tabprodutos.Open "Produtos", conectar, adOpenKeyset, adLockOptimistic

If tabfornecedores.State = adStateOpen Then tabfornecedores.Close
tabfornecedores.Open "Fornecedores", conectar, adOpenKeyset, adLockOptimistic

If tabpedidodecompra.State = adStateOpen Then tabpedidodecompra.Close
tabpedidodecompra.Open "Pedidos", conectar, adOpenKeyset, adLockOptimistic

If tabsegmentos.State = adStateOpen Then tabsegmentos.Close
tabsegmentos.Open "Segmentos", conectar, adOpenKeyset, adLockOptimistic

If tabmarca.State = adStateOpen Then tabmarca.Close
tabmarca.Open "Marcas", conectar, adOpenKeyset, adLockOptimistic

Call cod_pdc
dtp_data.value = Date$
Call flex
End Sub

Private Sub mfg_produtos_Click()

            If mfg_produtos.Rows < 2 Then Exit Sub
            L_linha = mfg_produtos.Row
            l_cod_pedido_compra = mfg_produtos.TextMatrix(L_linha, 0)
            If tabpedidodecompra.State = adStateOpen Then tabpedidodecompra.Close
            tabpedidodecompra.Open "Select * From Pedidos Where Codigo = " & l_cod_pedido_compra
            txt_np = tabpedidodecompra!codigo
            txt_codproduto = tabpedidodecompra!cod_produtos
            txt_codfornecedor = tabpedidodecompra!cod_fornecedor
            dtp_data.value = tabpedidodecompra!Data
            txt_codproduto_LostFocus
            txt_codfornecedor_LostFocus
            txt_qtde = tabpedidodecompra!Quantidade
            txt_total = Format(Round(tabpedidodecompra!Total), "currency")
            txt_codproduto.SetFocus
End Sub

Private Sub msk_cnpj_LostFocus()
Call desativar
If msk_cnpj <> "" Then

If tabfornecedores.State = adStateOpen Then tabfornecedores.Close
tabfornecedores.Open "select * from Fornecedores where CNPJ = '" & msk_cnpj & "'"

If tabfornecedores.RecordCount = 1 Then
txt_rs = tabfornecedores!Razao_social
msk_conta = tabfornecedores!Conta_corrente
txt_codfornecedor = tabfornecedores!codigo
msk_ie = tabfornecedores!Ie
End If

End If
Call ativar
End Sub

Private Sub msk_conta_LostFocus()
Call desativar
If msk_conta <> "" Then

If tabfornecedores.State = adStateOpen Then tabfornecedores.Close
tabfornecedores.Open "select * from Fornecedores where Conta_corrente = '" & msk_conta & "'"

If tabfornecedores.RecordCount = 1 Then
txt_rs = tabfornecedores!Razao_social
msk_ie = tabfornecedores!Ie
txt_codfornecedor = tabfornecedores!codigo
msk_cnpj = tabfornecedores!Cnpj
End If

End If
Call ativar
End Sub

Private Sub msk_ie_LostFocus()
Call desativar
If msk_ie <> "" Then

If tabfornecedores.State = adStateOpen Then tabfornecedores.Close
tabfornecedores.Open "select * from Fornecedores where IE = '" & msk_ie & "'"

If tabfornecedores.RecordCount = 1 Then
txt_rs = tabfornecedores!Razao_social
msk_conta = tabfornecedores!Conta_corrente
txt_codfornecedor = tabfornecedores!codigo
msk_cnpj = tabfornecedores!Cnpj
End If

End If
Call ativar
End Sub

Private Sub txt_codfornecedor_LostFocus()
If txt_codfornecedor <> "" Then

If tabfornecedores.State = adStateOpen Then tabfornecedores.Close
tabfornecedores.Open "select * from Fornecedores where Codigo =" & txt_codfornecedor

Call desativar
If tabfornecedores.RecordCount = 1 Then
txt_rs = tabfornecedores!Razao_social
msk_conta = tabfornecedores!Conta_corrente
msk_ie = tabfornecedores!Ie
msk_cnpj = tabfornecedores!Cnpj
End If
Call ativar

End If
End Sub

Private Sub txt_codproduto_LostFocus()
If txt_codproduto <> "" Then
If tabprodutos.State = adStateOpen Then tabprodutos.Close
tabprodutos.Open "select * from Produtos where Codigo =" & txt_codproduto

If tabprodutos.RecordCount <> 0 Then

If tabsegmentos.State = adStateOpen Then tabsegmentos.Close: tabsegmentos.Open "select * from Segmentos where Codigo =" & tabprodutos!Segmento
If tabmarca.State = adStateOpen Then tabmarca.Close: tabmarca.Open "select * from marcas where Codigo =" & tabprodutos!Marca

txt_nome = tabprodutos!nome
txt_unidade = tabprodutos!unidade
txt_segmento = tabsegmentos!Segmento
txt_marca = tabmarca!Marca
txt_vu = Format(tabprodutos!Preco_unitario, "currency")
Else
txt_segmento = Clear
txt_marca = Clear
txt_unidade = Clear
End If
End If
End Sub

Private Sub txt_nome_LostFocus()
If txt_nome <> "" Then
If tabprodutos.State = adStateOpen Then tabprodutos.Close
tabprodutos.Open "select * from Produtos where Nome = '" & txt_nome & "'"

If tabprodutos.RecordCount = 1 Then

If tabsegmentos.State = adStateOpen Then tabsegmentos.Close: tabsegmentos.Open "select * from Segmentos where Codigo =" & tabprodutos!Segmento
If tabmarca.State = adStateOpen Then tabmarca.Close: tabmarca.Open "select * from marcas where Codigo =" & tabprodutos!Marca

txt_codproduto = tabprodutos!codigo
txt_unidade = tabprodutos!unidade
txt_segmento = tabsegmentos!Segmento
txt_marca = tabmarca!Marca
txt_vu = Format(tabprodutos!Preco_unitario, "currency")
Else
txt_segmento = Clear
txt_marca = Clear
txt_unidade = Clear
End If
End If
End Sub

Private Sub desativar()
msk_conta.PromptInclude = False
msk_ie.PromptInclude = False
msk_cnpj.PromptInclude = False
End Sub

Private Sub ativar()
msk_conta.PromptInclude = True
msk_ie.PromptInclude = True
msk_cnpj.PromptInclude = True
End Sub

Private Sub txt_qtde_Change()
If txt_qtde = "" Then txt_total = "": Exit Sub
txt_total = Format(Round(txt_qtde * txt_vu), "currency")
End Sub

Private Sub txt_rs_LostFocus()
If txt_rs <> "" Then

If tabfornecedores.State = adStateOpen Then tabfornecedores.Close
tabfornecedores.Open "select * from Fornecedores where Razao_social = '" & txt_rs & "'"

Call desativar
If tabfornecedores.RecordCount = 1 Then
txt_codfornecedor = tabfornecedores!codigo
msk_conta = tabfornecedores!Conta_corrente
msk_ie = tabfornecedores!Ie
msk_cnpj = tabfornecedores!Cnpj
Else
txt_codfornecedor = Clear
msk_conta = Clear
msk_ie = Clear
msk_cnpj = Clear
End If
Call ativar

End If
End Sub

Private Sub cod_pdc()

            l_cod_pedido_compra = 1
A:
            If tabpedidodecompra.State = adStateOpen Then tabpedidodecompra.Close
            tabpedidodecompra.Open "select * from Pedidos where Codigo = " & l_cod_pedido_compra
            If tabpedidodecompra.RecordCount > 0 Then
                l_cod_pedido_compra = l_cod_pedido_compra + 1
                GoTo A:
            End If
            txt_np = l_cod_pedido_compra

End Sub

Private Sub gravar_pdc()
            If status <> "alteradas" Then
                If tabpedidodecompra.State = adStateOpen Then tabpedidodecompra.Close
                tabpedidodecompra.Open "Pedidos", conectar, adOpenKeyset, adLockOptimistic
                tabpedidodecompra.AddNew
                tabpedidodecompra!codigo = txt_np
            Else
                If tabpedidodecompra.State = adStateOpen Then tabpedidodecompra.Close
                tabpedidodecompra.Open "Select * from Pedidos where Codigo =" & txt_np
                If tabpedidodecompra.RecordCount = 0 Then
                MsgBox "PEDIDO NÃO CADASTRADO", vbInformation, "TRICON SUPERMERCADOS LTDA."
                var = 1
                Exit Sub
                End If
            End If
                tabpedidodecompra!Data = dtp_data.value
                tabpedidodecompra!Quantidade = txt_qtde
                tabpedidodecompra!Total = txt_total
                tabpedidodecompra!cod_produtos = txt_codproduto
                tabpedidodecompra!cod_fornecedor = txt_codfornecedor
                tabpedidodecompra.Update
End Sub
Private Sub flex()
            If tabpedidodecompra.State = adStateOpen Then tabpedidodecompra.Close
            tabpedidodecompra.Open "Pedidos", conectar, adOpenKeyset, adLockOptimistic
            If tabpedidodecompra!codigo <> "" Then tabpedidodecompra.MoveFirst
            mfg_produtos.Clear
            mfg_produtos.Rows = 2
            mfg_produtos.FormatString = "Numero do Pedido|Produto                          |Fornecedor                            |Quantidade    |Data            | Preço               "
            Do Until tabpedidodecompra.EOF = True
            
            If tabprodutos.State = adStateOpen Then tabprodutos.Close
            If tabfornecedores.State = adStateOpen Then tabfornecedores.Close
            tabprodutos.Open "Select * from Produtos where Codigo =" & tabpedidodecompra!cod_produtos
            tabfornecedores.Open "Select * from Fornecedores where Codigo =" & tabpedidodecompra!cod_fornecedor
            mfg_produtos.TextMatrix(mfg_produtos.Rows - 1, 0) = tabpedidodecompra!codigo
            mfg_produtos.TextMatrix(mfg_produtos.Rows - 1, 1) = tabprodutos!nome
            mfg_produtos.TextMatrix(mfg_produtos.Rows - 1, 2) = tabfornecedores!Razao_social
            mfg_produtos.TextMatrix(mfg_produtos.Rows - 1, 3) = tabpedidodecompra!Quantidade
            mfg_produtos.TextMatrix(mfg_produtos.Rows - 1, 4) = tabpedidodecompra!Data
            mfg_produtos.TextMatrix(mfg_produtos.Rows - 1, 5) = Format(tabpedidodecompra!Total, "currency")
            mfg_produtos.Rows = mfg_produtos.Rows + 1
            tabpedidodecompra.MoveNext
            
            Loop
            mfg_produtos.Rows = mfg_produtos.Rows - 1

End Sub
