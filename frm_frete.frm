VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_frete 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preço do frete"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5250
   Icon            =   "frm_frete.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_codigo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   22
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   2655
      Left            =   0
      TabIndex        =   19
      Top             =   2640
      Width           =   5175
      Begin MSFlexGridLib.MSFlexGrid mfg_fretes 
         Height          =   2295
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4048
         _Version        =   393216
         BackColorFixed  =   14737632
         BackColorBkg    =   12632256
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   0
      TabIndex        =   16
      Top             =   1920
      Width           =   5175
      Begin VB.TextBox txt_preco 
         Height          =   375
         Left            =   1560
         TabIndex        =   18
         Text            =   "R$"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Preço frete"
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
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1695
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   5175
      Begin VB.ComboBox cbo_bairro 
         Height          =   315
         Left            =   1560
         TabIndex        =   14
         Top             =   480
         Width           =   3375
      End
      Begin VB.ComboBox cbo_cidade 
         Height          =   315
         Left            =   3600
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
      Begin VB.ComboBox cbo_estado 
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
      Begin VB.ComboBox cbo_logradouro 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   840
         Width           =   3375
      End
      Begin MSMask.MaskEdBox msk_cep 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "99999-999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bairro"
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
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Estado"
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
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cidade"
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
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Logradouro"
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
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cep"
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
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Comandos"
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   5175
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
         Left            =   4080
         Picture         =   "frm_frete.frx":1CCA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Excluir"
         Top             =   240
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
         Left            =   2880
         Picture         =   "frm_frete.frx":1FD4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Alterar"
         Top             =   240
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
         Left            =   1680
         Picture         =   "frm_frete.frx":22DE
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salvar"
         Top             =   240
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
         Left            =   360
         Picture         =   "frm_frete.frx":25E8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Novo"
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label7 
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
      Left            =   240
      TabIndex        =   21
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "frm_frete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L_Colunas, L_Linha As Long
Dim L_Codfrete As Integer
Public Codfrete As New ADODB.Recordset
Public logradouro, bairro, cidade, estado As Integer
Dim c_e_p, L, B, C, E As String
Private Sub carregar_combos()
                Call abrir
                
            Do While tabuf.EOF = False
                    cbo_estado.AddItem tabuf!estado
                    cbo_estado.ItemData(cbo_estado.NewIndex) = tabuf!Codigo
                    tabuf.MoveNext
            Loop
            
            Do While tabcid.EOF = False
                    cbo_cidade.AddItem tabcid!cidade
                    cbo_cidade.ItemData(cbo_cidade.NewIndex) = tabcid!Cod_cidade
                    tabcid.MoveNext
            Loop
            
            Do While tabbairro.EOF = False
                    cbo_bairro.AddItem tabbairro!bairro
                    cbo_bairro.ItemData(cbo_bairro.NewIndex) = tabbairro!Cod_bairro
                    tabbairro.MoveNext
            Loop
            
            Do While tablocalizacao.EOF = False
                    cbo_logradouro.AddItem tablocalizacao!logradouro
                   ' cbo_logradouro.ItemData(cbo_logradouro.NewIndex) = tablocalizacao!Logradouro
                    tablocalizacao.MoveNext
            Loop
End Sub



Private Sub cbo_logradouro_LostFocus()
           ' Call desabilitar_mascara
            'Call abrir
            'If cbo_logradouro <> Empty Then
            '    tablocalizacao.Close
            '    tablocalizacao.Open "select Logradouro from Localizacoes where Logradouro like '" & cbo_logradouro.Text & "'"
            '    If tablocalizacao.RecordCount = 1 Then
                
            '    End If
          '  End If
          '  Call habilitar_mascara
End Sub

Private Sub cmd_alterar_Click()
             status = "alteradas"
            If tabfrete.State = adStateOpen Then tabfrete.Close
                tabfrete.Open "Select * From Fretes where Frete like '" & txt_codigo & "'"
                If tabfrete.RecordCount <> 0 Then
                Call alterar
                End If
            Call box
            Call carregar_lista
End Sub

Private Sub cmd_excluir_Click()
            status = "excluidas"
            If MsgBox("Deseja realmente excluir estas informações?", vbYesNo + vbDefaultButton2 + vbQuestion, "Arbimy manager 2.0") = vbYes Then
                If tabfrete.State = adStateOpen Then tabfrete.Close
                    tabfrete.Open "select * from Fretes where Frete = " & txt_codigo
                    If tabfrete.RecordCount = 1 Then
                        conectar.Execute "Delete From Fretes where Frete like '" & txt_codigo & "'"
              Call box
              Call carregar_lista
              Call limpar
              cmd_alterar.Enabled = False
              cmd_excluir.Enabled = False
                End If
                    End If
End Sub

Private Sub cmd_novo_Click()
            Call limpar_frete
            Call frete
            msk_cep.SetFocus
            
End Sub

Private Sub comd_salvar_Click()
            status = "gravadas"
            
            If txt_frete = Empty Then
                MsgBox "Descupe-nos! Há informações que não foram preenchidas", vbExclamation, "Tricon supermercados LTDA."
                End If
                Call Gravar_frete
                Call box
                Call limpar_frete
                Call fechar
                Call abrir
                Call carregar_lista
End Sub

Private Sub cmd_salvar_Click()
            status = "gravadas"
            If msk_cep = Empty Then
                MsgBox "Descupe-nos! Há informações que não foram preenchidas", vbExclamation, "Arbimy manager 2.0"
                Exit Sub
            End If
                Call gravar
                Call box
                Call limpar
                Call carregar_lista
End Sub

Private Sub Form_Load()
            carregar_combos
            Call frete
            Call carregar_lista
End Sub
Private Sub abrir()
            If tabuf.State = adStateOpen Then tabuf.Close
                tabuf.Open "Ufs", conectar, adOpenKeyset, adLockOptimistic
            
            If tabcid.State = adStateOpen Then tabcid.Close
                tabcid.Open "Cidades", conectar, adOpenKeyset, adLockOptimistic
            
            If tabbairro.State = adStateOpen Then tabbairro.Close
                tabbairro.Open "Bairros", conectar, adOpenKeyset, adLockOptimistic
            
            If tablocalizacao.State = adStateOpen Then tablocalizacao.Close
                tablocalizacao.Open "Localizacoes", conectar, adOpenKeyset, adLockOptimistic
            
            If Codfrete.State = adStateOpen Then Codfrete.Close
                Codfrete.Open "Fretes", conectar, adOpenKeyset, adLockOptimistic
            If tabfrete.State = adStateOpen Then tabfrete.Close
                tabfrete.Open "Fretes", conectar, adOpenKeyset, adLockOptimistic
End Sub
Private Sub habilitar_mascara()
            msk_cep.PromptInclude = True
End Sub
Private Sub desabilitar_mascara()
            msk_cep.PromptInclude = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
            Call voltar_botao
End Sub

Private Sub mfg_fretes_Click()
            L_Linha = mfg_fretes.Row
            L_Codfrete = mfg_fretes.TextMatrix(L_Linha, 0)
            Call abrir
            tabfrete.Close
            tabfrete.Open "Select * From Fretes Where Frete = " & L_Codfrete
            Call Mostrar
            cmd_alterar.Enabled = True
            cmd_excluir.Enabled = True
            cbo_logradouro.Enabled = False
            cbo_bairro.Enabled = False
            cbo_cidade.Enabled = False
            cbo_estado.Enabled = False
            Call carregar_lista
            
End Sub

Private Sub msk_cep_LostFocus()
                Dim cod_b As Integer
                Dim cod_c As Integer
            
            Call desabilitar_mascara
            
            If tablocalizacao.State = adStateOpen Then tablocalizacao.Close
            tablocalizacao.Open "select * from Localizacoes where Cep = '" & msk_cep & "'"
            If tablocalizacao.RecordCount = 1 Then
            cbo_logradouro.Text = tablocalizacao!logradouro
            cod_b = tablocalizacao!Cod_bairro
            End If
            If tablocalizacao.RecordCount = 0 Then
            If MsgBox("Esta localização ainda não está cadastrado! Deseja cadastrar agora?", vbYesNo + vbDefaultButton1 + vbQuestion, "Arbimy manager 2.0") = vbYes Then frm_localizacoes.Show
            Exit Sub
            msk_cep.Text = Empty
            Exit Sub
            Else
            End If
            If tabbairro.State = adStateOpen Then tabbairro.Close
            tabbairro.Open "select * from Bairros where Cod_bairro = " & cod_b
            If tabbairro.RecordCount = 1 Then
            cod_c = tabbairro!Cod_cidade
            cbo_bairro.Text = tabbairro!bairro
            End If
            If tabcid.State = adStateOpen Then tabcid.Close
            tabcid.Open "select * from Cidades where Cod_cidade = " & cod_c
            If tabcid.RecordCount = 1 Then
            cod_u = tabcid!Cod_estado
            cbo_cidade.Text = tabcid!cidade
            End If
            If tabuf.State = adStateOpen Then tabuf.Close
            tabuf.Open "select * from Ufs where Codigo = " & cod_u
            If tabuf.RecordCount = 1 Then
            cbo_estado.Text = tabuf!estado
            Call habilitar_mascara
            End If
            cbo_estado.Enabled = False
            cbo_cidade.Enabled = False
            cbo_bairro.Enabled = False
            cbo_logradouro.Enabled = False
End Sub
Private Sub txt_preco_GotFocus()
            If txt_preco.Text = "R$" Then
                txt_preco.Text = Clear
            End If
End Sub
Private Sub txt_preco_LostFocus()
            txt_preco = Format(txt_preco, "currency")
End Sub
Private Sub limpar_frete()
            Call desabilitar_mascara
            cbo_estado.Text = Clear
            cbo_cidade.Text = Clear
            cbo_bairro.Text = Clear
            cbo_logradouro.Text = Clear
            msk_cep.Text = Clear
            txt_frete = Clear
            txt_preco = Clear
            Call habilitar_mascara
End Sub
Private Sub Gravar_frete()
            Call desabilitar_mascara
            
            If status <> "alteradas" Then tabfrete.AddNew
            tabfrete!cep = msk_cep.Text
            tabfrete!Preco = txt_preco.Text
            Call habilitar_mascara
            tabfrete.Update
End Sub
Private Sub fechar()
            If tabfrete.State = adStateOpen Then tabfrete.Close
            If Codfrete.State = adStateOpen Then Codfrete.Close
End Sub
Private Sub frete()
            Dim codfrete2 As Integer
            codfrete2 = 1
A:
            Call fechar
            Codfrete.Open "Fretes", conectar, adOpenKeyset, adLockOptimistic
            Call fechar
            Codfrete.Open "select * from fretes where Frete = " & codfrete2
            If Codfrete.RecordCount > 0 Then
                codfrete2 = codfrete2 + 1
                GoTo A
            End If
            txt_codigo = codfrete2
End Sub
Private Sub carregar_lista()
            Call abrir
            If tabfrete.State = adStateOpen Then tabfrete.Close
            tabfrete.Open "select * from fretes"
            If tabfrete.RecordCount = 0 Then
                MsgBox "Descupe-nos, não temos nenhum frete cadastrado", vbInformation, "Arbimy manager 2.0"
                Exit Sub
            End If
                mfg_fretes.Rows = 2
                mfg_fretes = Clear
                mfg_fretes.FormatString = "Código  |Logradouro                                              |Preço do frete "
            tabfrete.MoveFirst
                'Call converter
            Do While tabfrete.EOF = False
                'c_e_p = tabfrete!Cep
                mfg_fretes.TextMatrix(mfg_fretes.Rows - 1, 0) = tabfrete!frete
                'mfg_fretes.TextMatrix(mfg_fretes.Rows - 1, 1) = L
                mfg_fretes.TextMatrix(mfg_fretes.Rows - 1, 2) = Format(tabfrete!Preco, "currency")
                tabfrete.MoveNext
                mfg_fretes.Rows = mfg_fretes.Rows + 1
            Loop
                mfg_fretes.Rows = mfg_fretes.Rows - 1
                
End Sub
Private Sub Mostrar()
            If txt_preco.Text = "R$" Then
                txt_preco.Text = Clear
                msk_cep.SetFocus
            End If
            Call desabilitar_mascara
            txt_codigo.Text = tabfrete!frete
            msk_cep.Text = tabfrete!cep
            txt_preco = Format(tabfrete!Preco, "currency")
            'Call converter
            'cbo_logradouro = L
            cmd_alterar.Enabled = True
            cmd_excluir.Enabled = True
            Call habilitar_mascara
            'Call converter2
            
End Sub
Private Sub converter()
            Call abrir
            Codfrete.Close
            Codfrete.Open "select * from Fretes where Cep like '" & c_e_p & "'"
            If Codfrete.RecordCount > 0 Then
            tablocalizacao.Close
            tablocalizacao.Open "Select * from  Localizacoes where Cep = '" & c_e_p & "'"
            If tablocalizacao.RecordCount = 1 Then
            L = tablocalizacao!logradouro
            End If
            End If
End Sub
Private Sub gravar()
            Call desabilitar_mascara
            Call abrir
            tabfrete.Close
            tabfrete.Open "select * from Fretes where Cep like " & msk_cep
            If tabfrete.RecordCount = 0 Then
            If status <> "alteradas" Then Codfrete.AddNew
            Codfrete!frete = txt_codigo.Text
            Codfrete!Preco = txt_preco.Text
            Codfrete!cep = msk_cep.Text
            Codfrete.Update
            End If
            Call habilitar_mascara
End Sub
Private Sub limpar()
            Call desabilitar_mascara
            txt_codigo.Text = Clear
            cbo_estado.Text = Clear
            cbo_cidade.Text = Clear
            cbo_bairro.Text = Clear
            cbo_logradouro.Text = Clear
            msk_cep.Text = Clear
            txt_preco.Text = Clear
            msk_cep.SetFocus
            Call habilitar_mascara
End Sub
Private Sub alterar()
            Call desabilitar_mascara
            tabfrete!cep = msk_cep.Text
            tabfrete!frete = txt_codigo.Text
            tabfrete!Preco = txt_preco.Text
            tabfrete.Update
            Call habilitar_mascara
End Sub
Private Sub converter2()
            Call desabilitar_mascara
            c_e_p = msk_cep.Text
            If tabfrete.State = adStateOpen Then
            tabfrete.Close
            tabfrete.Open "select * from Fretes where Cep = '" & c_e_p & "'"
            End If
            
            If tabfrete.RecordCount = 1 Then
            If tablocalizacao.State = adStateOpen Then tablocalizacao.Close
            tablocalizacao.Open "select* from Localizacoes where Cep = '" & c_e_p & "'"
            End If
            
            If tablocalizacao.RecordCount = 1 Then
            bairro = tablocalizacao!Cod_bairro
            End If
            
            If tabbairro.State = adStateOpen Then tabbairro.Close
            tabbairro.Open "select * from Bairros where Cod_bairro like '" & bairro & "'"
            If tabbairro.RecordCount = 1 Then
            cbo_bairro = tabbairro!bairro
            cidade = tabbairro!Cod_cidade
            End If
            
            If tabcid.State = adStateOpen Then tabcid.Close
            tabcid.Open "select * from Cidades where Cod_cidade like '" & cidade & "'"
            If tabcid.RecordCount = 1 Then
            cbo_cidade = tabcid!cidade
            estado = tabcid!Cod_estado
            End If
            
            If tabuf.State = adStateOpen Then tabuf.Close
            tabuf.Open "select * from Ufs where Codigo like '" & estado & "'"
            If tabcid.RecordCount = 1 Then
            cbo_estado = tabuf!estado
            End If
            
            
            
            
            
            
End Sub
