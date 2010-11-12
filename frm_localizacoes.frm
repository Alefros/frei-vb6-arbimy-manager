VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_localizacoes 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Localizações"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5115
   Icon            =   "frm_localizacoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frm_localizacoes.frx":030A
   ScaleHeight     =   6585
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid mfg_localizacao 
      Height          =   2295
      Left            =   0
      TabIndex        =   21
      Top             =   4200
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4048
      _Version        =   393216
      BackColorFixed  =   14737632
      BackColorSel    =   12632256
      BackColorBkg    =   12632256
   End
   Begin VB.ComboBox cbo_localizacao 
      Height          =   315
      ItemData        =   "frm_localizacoes.frx":0614
      Left            =   0
      List            =   "frm_localizacoes.frx":0624
      TabIndex        =   20
      Text            =   "(Selecione aqui qual informação deseja buscar)"
      Top             =   3840
      Width           =   3735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Comandos"
      ForeColor       =   &H80000006&
      Height          =   1095
      Left            =   0
      TabIndex        =   19
      Top             =   2640
      Width           =   5055
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
         Picture         =   "frm_localizacoes.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
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
         Left            =   3960
         Picture         =   "frm_localizacoes.frx":0956
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   2760
         Picture         =   "frm_localizacoes.frx":0C60
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   1560
         Picture         =   "frm_localizacoes.frx":0F6A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2655
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "Estados"
      TabPicture(0)   =   "frm_localizacoes.frx":1274
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_estado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txt_estado"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Cidades"
      TabPicture(1)   =   "frm_localizacoes.frx":1290
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cbo_estado"
      Tab(1).Control(1)=   "txt_cidade"
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(3)=   "Label2"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Bairros"
      TabPicture(2)   =   "frm_localizacoes.frx":12AC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cbo_cidade"
      Tab(2).Control(1)=   "txt_bairro"
      Tab(2).Control(2)=   "Label6"
      Tab(2).Control(3)=   "Label3"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Localizações"
      TabPicture(3)   =   "frm_localizacoes.frx":12C8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cbo_bairro"
      Tab(3).Control(1)=   "txt_logradouro"
      Tab(3).Control(2)=   "msk_cep"
      Tab(3).Control(3)=   "Label8"
      Tab(3).Control(4)=   "Label7"
      Tab(3).Control(5)=   "Label4"
      Tab(3).ControlCount=   6
      Begin VB.ComboBox cbo_bairro 
         Height          =   315
         Left            =   -73680
         TabIndex        =   18
         Top             =   1440
         Width           =   2535
      End
      Begin VB.ComboBox cbo_cidade 
         Height          =   315
         Left            =   -73680
         TabIndex        =   5
         Top             =   1440
         Width           =   2535
      End
      Begin VB.ComboBox cbo_estado 
         Height          =   315
         Left            =   -73680
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txt_logradouro 
         Height          =   285
         Left            =   -73680
         TabIndex        =   14
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txt_bairro 
         Height          =   285
         Left            =   -73680
         TabIndex        =   4
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txt_cidade 
         Height          =   285
         Left            =   -73680
         TabIndex        =   2
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txt_estado 
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   1
         Top             =   1080
         Width           =   1215
      End
      Begin MSMask.MaskEdBox msk_cep 
         Height          =   375
         Left            =   -73680
         TabIndex        =   22
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "99999-999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         Caption         =   "Cep"
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
         Left            =   -74880
         TabIndex        =   23
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Bairro"
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
         Left            =   -74880
         TabIndex        =   17
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Cidade"
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
         Left            =   -74880
         TabIndex        =   16
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Estado"
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
         Left            =   -74880
         TabIndex        =   15
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Logradouro"
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
         Left            =   -74880
         TabIndex        =   13
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Bairro"
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
         Left            =   -74880
         TabIndex        =   12
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Cidade"
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
         Left            =   -74880
         TabIndex        =   11
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lbl_estado 
         Caption         =   "Estado"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   735
      End
   End
End
Attribute VB_Name = "frm_localizacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L_cod_uf As Integer
Private Sub cbo_estado_Change()
           ' L_cod_uf = cbo_estados.ItemData(cbo_estados.ListIndex)
End Sub

Private Sub cbo_localizacao_Click()
If cbo_localizacao.Text = "Estado" Then
                MsgBox "estado"
                Exit Sub
            
            ElseIf cbo_localizacao.Text = "Cidade" Then
                MsgBox "Cidade"
                Exit Sub
                
                ElseIf cbo_localizacao.Text = "Bairro" Then
                MsgBox "bairro"
                Exit Sub
                    
                    ElseIf cbo_localizacao.Text = "Logradouro" Then
                    MsgBox "logradouro"
                    Exit Sub
            End If
End Sub

Private Sub cmd_alterar_Click()
               status = "alteradas"
            
            If SSTab1.Tab = 0 Then
               status = "alteradas"
                    If tabuf.State = adStateOpen Then tabuf.Close
                        tabuf.Open "Select * From Ufs where Estado like '" & txt_estado & "'"
                    If tabuf.RecordCount <> 0 Then
                        'Call alterar
             '       End If
            'Call box
            
                ElseIf SSTab1.Tab = 1 Then
                   'status = "alteradas"
            'If tabcli.State = adStateOpen Then tabcli.Close
             '   tabcli.Open "Select * From Clientes where Codigo like '" & txt_codigo & "'"
              '  If tabcli.RecordCount <> 0 Then
               ' Call alterar
                'End If
            'Call box
            
                    ElseIf SSTab1.Tab = 2 Then
                    
                        ElseIf SSTab1.Tab = 3 Then
                
            End If
            End If
End Sub

Private Sub cmd_excluir_Click()
            If SSTab1.Tab = 0 Then
            
                ElseIf SSTab1.Tab = 1 Then
                
                    ElseIf SSTab1.Tab = 2 Then
    
                        ElseIf SSTab1.Tab = 3 Then
                        
            End If
End Sub

Private Sub cmd_novo_Click()
            
            If SSTab1.Tab = 0 Then 'estados
            Call limpar_localizacao
             txt_estado.SetFocus
                
                ElseIf SSTab1.Tab = 1 Then 'cidades
                Call limpar_localizacao
                 txt_cidade.SetFocus
                    
                    ElseIf SSTab1.Tab = 2 Then ' bairros
                    Call limpar_localizacao
                     txt_bairro.SetFocus
                        
                        ElseIf SSTab1.Tab = 3 Then 'localizações
                        Call limpar_localizacao
                         txt_logradouro.SetFocus
            End If
End Sub

Private Sub cmd_salvar_Click()
            status = "gravadas"
            
            If SSTab1.Tab = 0 Then 'estados
                If txt_estado = Empty Then
                MsgBox "Descupe-nos! Há informações que não foram preenchidas", vbExclamation, "Tricon supermercados LTDA."
                End If
                Call gravar_estado
                Call box
                Call limpar_localizacao
                
                ElseIf SSTab1.Tab = 1 Then 'cidades
                        If txt_cidade = Empty Then
                        MsgBox "Descupe-nos! Há informações que não foram preenchidas", vbExclamation, "Tricon supermercados LTDA."
                        End If
                        Call gravar_cidade
                        Call box
                        Call limpar_localizacao
                        
                    ElseIf SSTab1.Tab = 2 Then 'bairros
                            If txt_bairro = Empty Then
                            MsgBox "Descupe-nos! Há informações que não foram preenchidas", vbExclamation, "Tricon supermercados LTDA."
                            End If
                            Call gravar_bairro
                            Call box
                            Call limpar_localizacao
                            
                        ElseIf SSTab1.Tab = 3 Then 'localizações
                                If txt_logradouro = Empty Then
                                MsgBox "Descupe-nos! Há informações que não foram preenchidas", vbExclamation, "Tricon supermercados LTDA."
                                End If
                                Call gravar_logradouro
                                Call box
                                Call limpar_localizacao
            End If
                    Call fechar
                    Call abrir
                    Call carregar_combos
End Sub
Private Sub limpar_localizacao()
            Call desabilitar_mascara
            msk_cep = Clear
            txt_bairro = Clear
            txt_cidade = Clear
            txt_estado = Clear
            txt_logradouro = Clear
            cbo_estado = Clear
            cbo_cidade = Clear
            cbo_bairro = Clear
            Call habilitar_mascara
End Sub
Private Sub gravar_estado()
            If status <> "alteradas" Then tabuf.AddNew
            tabuf!estado = txt_estado.Text
            tabuf.Update
End Sub
Private Sub gravar_cidade()
            If status <> "alteradas" Then tabcid.AddNew
            tabcid!cidade = txt_cidade.Text
            If tabuf.State = adStateOpen Then tabuf.Close
            tabuf.Open "select * from Ufs where Estado = '" & cbo_estado & "'"           'ABRA A CONEXÃO COM A TAB_UF E SELECIONE TODOS CAMPOS DA TABELA UF ONDE NOMES SERÁ ATRIBUIDO DA TXT_ESTADO
            If tabuf.RecordCount <> 0 Then
            tabcid!Cod_estado = tabuf!Codigo
            End If
            tabcid.Update
End Sub
Private Sub gravar_bairro()
            If status <> "alteradas" Then tabbairro.AddNew
            tabbairro!bairro = txt_bairro.Text
            If tabcid.State = adStateOpen Then tabcid.Close
            tabcid.Open "select * from Cidades where Cidade = '" & cbo_cidade & "'"
            If tabcid.RecordCount <> 0 Then
            tabbairro!Cod_cidade = tabcid!Cod_cidade
            End If
            tabbairro.Update
End Sub
Private Sub gravar_logradouro()
            Call desabilitar_mascara
            If status <> "alteradas" Then tablocalizacao.AddNew
            tablocalizacao!logradouro = txt_logradouro.Text
            If tabbairro.State = adStateOpen Then tabbairro.Close
            tabbairro.Open "select * from Bairros where Bairro = '" & cbo_bairro & "'"
            If tabbairro.RecordCount <> 0 Then
            tablocalizacao!Cod_bairro = tabbairro!Cod_bairro
            tablocalizacao!cep = msk_cep.Text
            End If
            Call habilitar_mascara
            tablocalizacao.Update
End Sub
Private Sub desabilitar_mascara()
            msk_cep.PromptInclude = False
End Sub
Private Sub habilitar_mascara()
            msk_cep.PromptInclude = True
End Sub
Private Sub Form_Load()
            Call abrir
            Call carregar_combos
End Sub

Private Sub Form_Unload(Cancel As Integer)
            Call voltar_botao
            
End Sub

Private Sub txt_bairro_LostFocus()
            txt_bairro = UCase(txt_bairro)
End Sub

Private Sub txt_cidade_LostFocus()
            txt_cidade = UCase(txt_cidade)
End Sub

Private Sub txt_estado_LostFocus()
            txt_estado = UCase(txt_estado)
End Sub
Private Sub txt_logradouro_LostFocus()
            txt_logradouro = UCase(txt_logradouro)
End Sub

Private Sub carregar_combos()
    
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
            
End Sub
Private Sub fechar()
            If tablocalizacao.State = adStateOpen Then tablocalizacao.Close
End Sub
