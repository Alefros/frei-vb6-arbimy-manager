VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_clientes 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de clientes"
   ClientHeight    =   8115
   ClientLeft      =   3840
   ClientTop       =   525
   ClientWidth     =   7995
   Icon            =   "frm_clientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker dtp_nascimento 
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16580609
      CurrentDate     =   40466
   End
   Begin MSFlexGridLib.MSFlexGrid mfg_clientes 
      Height          =   3255
      Left            =   120
      TabIndex        =   36
      Top             =   4800
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5741
      _Version        =   393216
      BackColorFixed  =   14737632
      BackColorSel    =   12632256
      BackColorBkg    =   12632256
   End
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
      Height          =   4455
      Left            =   6960
      TabIndex        =   35
      Top             =   240
      Width           =   1095
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
         Picture         =   "frm_clientes.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmd_excluir 
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
         Left            =   120
         Picture         =   "frm_clientes.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3600
         Width           =   855
      End
      Begin VB.CommandButton cmd_alterar 
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
         Left            =   120
         Picture         =   "frm_clientes.frx":0EDE
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmd_salvar 
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
         Left            =   120
         Picture         =   "frm_clientes.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Endereço"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   27
      Top             =   2760
      Width           =   6735
      Begin VB.TextBox txt_complemento 
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txt_numero 
         Height          =   375
         Left            =   5280
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txt_logradouro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         MaxLength       =   35
         TabIndex        =   13
         Top             =   1440
         Width           =   5055
      End
      Begin VB.TextBox txt_bairro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3840
         MaxLength       =   35
         TabIndex        =   12
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txt_cidade 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         MaxLength       =   35
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txt_uf 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin MSMask.MaskEdBox msk_cep 
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "99999-999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Complemento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Numero *"
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
         Left            =   3960
         TabIndex        =   33
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label12 
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
         TabIndex        =   32
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label9 
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
         Left            =   3000
         TabIndex        =   31
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "UF"
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
         Left            =   3960
         TabIndex        =   30
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label10 
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
         Left            =   240
         TabIndex        =   29
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cep *"
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
         TabIndex        =   28
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Contato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   23
      Top             =   1440
      Width           =   6735
      Begin VB.TextBox txt_email 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3840
         MaxLength       =   35
         TabIndex        =   6
         Top             =   720
         Width           =   2775
      End
      Begin MSMask.MaskEdBox msk_telefone 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "(99)9999-9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_celular 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "(99)9999-9999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Telefone"
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
         TabIndex        =   26
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Email"
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
         Left            =   3000
         TabIndex        =   25
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Celular"
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
         TabIndex        =   24
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.TextBox txt_codigo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   19
      Top             =   0
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informações pessoais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   6735
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   1560
         MaxLength       =   35
         TabIndex        =   3
         Top             =   840
         Width           =   5055
      End
      Begin MSMask.MaskEdBox msk_rg 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "99.999.999-&"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nascimento *"
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
         Left            =   3840
         TabIndex        =   22
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rg *"
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
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nome *"
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
         TabIndex        =   20
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Label Label15 
      BackColor       =   &H00E0E0E0&
      Caption         =   "(*) Preenchimento obrigatório"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   37
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label1 
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
      TabIndex        =   18
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frm_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L_Colunas, L_linha As Long
Dim cod_b, cod_c, cod_u As Integer
'Dim cod_c As Integer
'Dim cod_u As Integer
Dim codloca As Integer
Dim coduf As Integer
Dim codcid As Integer
Dim codbairro As Integer
Dim L_Codcli As Integer
Dim c_e_p, L As String
'Dim L As String
Public tabuf As New ADODB.Recordset
Public tabusuarios As New ADODB.Recordset
Public tabcid As New ADODB.Recordset
Public tabcli As New ADODB.Recordset
Public tabcli2 As New ADODB.Recordset
Public Codcli As New ADODB.Recordset
Private Sub cmd_alterar_Click()
                status = "alteradas"
            If tabcli.State = adStateOpen Then tabcli.Close
                tabcli.Open "Select * From Clientes where Codigo like '" & txt_codigo & "'"
                If tabcli.RecordCount <> 0 Then
                Call alterar
                End If
            Call box
End Sub
Private Sub cmd_excluir_Click()
              status = "excluidas"
            If MsgBox("Deseja realmente excluir este cliente?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
                If tabcli.State = adStateOpen Then tabcli.Close
                    tabcli.Open "select * from Clientes where codigo = " & txt_codigo
                    If tabcli.RecordCount <> 0 Then
                        conectar.Execute "Delete From Clientes where Codigo like '" & txt_codigo & "'"
              Call box
              Call carregar_lista
              Call limpar_cliente
              cmd_alterar.Enabled = False
              cmd_excluir.Enabled = False
                End If
                    End If
End Sub
Private Sub cmd_novo_Click()
            Call limpar_cliente
             msk_rg.SetFocus
             cmd_salvar.Enabled = False
             cmd_alterar.Enabled = False
             cmd_excluir.Enabled = False
End Sub
Private Sub limpar_cliente()
            Call desabilitar_mascara
            msk_telefone.Text = Clear
            msk_cep.Text = Clear
            msk_celular.Text = Clear
            msk_rg = Clear
              dtp_nascimento = Clear
             txt_codigo.Text = Clear
             txt_nome = Clear
             txt_codigo = Clear
             txt_email = Clear
             txt_cidade = Clear
             txt_uf = Clear
             txt_bairro = Clear
             txt_logradouro = Clear
             txt_numero = Clear
             txt_complemento = Clear
            Call habilitar_mascara
            Call cod_cli
End Sub
Private Sub desabilitar_mascara()
            msk_telefone.PromptInclude = False
            msk_rg.PromptInclude = False
            msk_cep.PromptInclude = False
            msk_celular.PromptInclude = False
End Sub
Private Sub habilitar_mascara()
            msk_telefone.PromptInclude = True
            msk_rg.PromptInclude = True
            msk_cep.PromptInclude = True
            msk_celular.PromptInclude = True
            dtp_nascimento.value = Date
End Sub
Private Sub cmd_salvar_Click()
                status = "gravadas"
            If txt_nome = Empty Or msk_rg = Empty Or msk_cep = Empty Or txt_numero = Empty Then
                MsgBox "Descupe-nos! Há informações que não foram preenchidas", vbExclamation, "Arbimy manager 2.0"
                Exit Sub
            End If
                Call gravar_cliente
                Call box
                Call limpar_cliente
                Call carregar_lista
End Sub
Private Sub gravar_cliente()
            Call desabilitar_mascara
            If tabcli.State = adStateOpen Then tabcli.Close
            tabcli.Open "select * from Clientes where Rg like " & msk_rg.Text
            If tabcli.RecordCount = 0 Then
            If status <> "alteradas" Then Codcli.AddNew
            Codcli!nome = txt_nome
            Codcli!Cep = msk_cep.Text
            Codcli!Email = txt_email
            Codcli!Numero = txt_numero
            Codcli!Complemento = txt_complemento
            Codcli!Nascimento = dtp_nascimento.value
            Codcli!Tel_res = msk_telefone.Text
            If msk_celular = Empty Then Codcli!Celular = "0"
            Else
            Codcli!Celular = msk_celular.Text
            
            Codcli!Rg = msk_rg
            Codcli!codigo = txt_codigo
            Codcli.Update
            Call cod_cli
            Call habilitar_mascara
            End If
End Sub
Private Sub Form_Load()
            If tabcli.State = adStateOpen Then tabcli.Close
                tabcli.Open "Clientes", conectar, adOpenKeyset, adLockOptimistic
            Call abrir
            Call carregar_lista
            Call cod_cli
End Sub
Private Sub cod_cli()
            Dim codcli2 As Integer
            codcli2 = 1
A:
            If Codcli.State = adStateOpen Then Codcli.Close
            Codcli.Open "Clientes", conectar, adOpenKeyset, adLockOptimistic
            If Codcli.State = adStateOpen Then Codcli.Close
            Codcli.Open "select * from Clientes where Codigo = " & codcli2
            If Codcli.RecordCount > 0 Then
                codcli2 = codcli2 + 1
                GoTo A
            End If
            txt_codigo = codcli2
End Sub
Private Sub bancos()
            tabcli!nome = txt_nome
            tabcli!Cep = msk_cep.Text
            tabcli!Email = txt_email
            tabcli!Numero = txt_numero
            tabcli!Complemento = txt_complemento
            tabcli!Nascimento = dtp_nascimento.value
            tabcli!Tel_res = msk_telefone.Text
            tabcli!Celular = msk_celular.Text
            tabcli!Rg = msk_rg
            tabcli!codigo = txt_codigo
            tabcli.Update
End Sub
Private Sub alterar()
            desabilitar_mascara
            Call bancos
            habilitar_mascara
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
            If tabcli2.State = adStateOpen Then tabcli2.Close
                tabcli2.Open "Clientes", conectar, adOpenKeyset, adLockOptimistic
            If tabusuarios.State = 1 Then tabusuarios.Close
            tabusuarios.Open "Usuarios", conectar, adOpenKeyset, adLockOptimistic
End Sub
Private Sub fechar()
            If tabuf.State = adStateOpen Then tabuf.Close
            If tabcid.State = adStateOpen Then tabcid.Close
            If tabbairro.State = adStateOpen Then tabbairro.Close
            If tablocalizacao.State = adStateOpen Then tablocalizacao.Close
            If tabcli.State = adStateOpen Then tabcli.Close
End Sub
Private Sub Form_Unload(Cancel As Integer)
           ' MDIForm1.Tbl_padrao.Buttons.Item(1).value = tbrUnpressed
            ' MDIForm1.Tbl_padrao.Buttons.Item(1).Style = tbrDropdown
End Sub
Private Sub mfg_clientes_Click()
            L_linha = mfg_clientes.Row
            L_Codcli = mfg_clientes.TextMatrix(L_linha, 0)
            Call abrir
            tabcli.Close
            tabcli.Open "Select * From Clientes Where Codigo = " & L_Codcli
            Call Mostrar
            cmd_alterar.Enabled = True
            cmd_excluir.Enabled = True
End Sub
Private Sub Mostrar()
                On Error Resume Next
            desabilitar_mascara
            txt_codigo.Text = tabcli!codigo
            msk_telefone.Text = tabcli!Tel_res
            msk_rg.Text = tabcli!Rg
            msk_cep.Text = tabcli!Cep
            If tablocalizacao.State = adStateOpen Then tablocalizacao.Close
            tablocalizacao.Open "select * from Localizacoes where Cep = '" & msk_cep & "'"
            If tablocalizacao.RecordCount = 1 Then
            txt_logradouro.Text = tablocalizacao!logradouro
            cod_b = tablocalizacao!Cod_bairro
            End If
            If tablocalizacao.RecordCount = 0 Then
            If MsgBox("Esta localização ainda não está cadastrado! Deseja cadastrar agora?", vbYesNo + vbDefaultButton1 + vbQuestion, "Arbimy manager 2.0") = vbYes Then frm_localizacoes.Show
            Exit Sub
            txt_bairro = Clear
            txt_cidade = Clear
            txt_logradouro = Clear
            txt_estado = Clear
            msk_cep.Text = Empty
            Exit Sub
            Else
            End If
            If tabbairro.State = adStateOpen Then tabbairro.Close
            tabbairro.Open "select * from Bairros where Cod_bairro = " & cod_b
            If tabbairro.RecordCount = 1 Then
            cod_c = tabbairro!Cod_cidade
            txt_bairro.Text = tabbairro!bairro
            End If
            If tabcid.State = adStateOpen Then tabcid.Close
            tabcid.Open "select * from Cidades where Cod_cidade = " & cod_c
            If tabcid.RecordCount = 1 Then
            cod_u = tabcid!Cod_estado
            txt_cidade.Text = tabcid!cidade
            End If
            If tabuf.State = adStateOpen Then tabuf.Close
            tabuf.Open "select * from Ufs where Codigo = " & cod_u
            If tabuf.RecordCount = 1 Then
            txt_uf.Text = tabuf!estado
              Call habilitar_mascara
            msk_celular.Text = tabcli!Celular
            dtp_nascimento.value = tabcli!Nascimento
            txt_nome.Text = tabcli!nome
            txt_email.Text = tabcli!Email
            habilitar_mascara
            End If
End Sub
Private Sub carregar_lista()
            Call abrir
            If tabcli.State = adStateOpen Then tabcli.Close
            tabcli.Open "select * from Clientes"
            If tabcli.RecordCount = 0 Then
                MsgBox "Descupe-nos, não temos nenhum cliente cadastrado", vbInformation, "Arbimy manager 2.0"
                Exit Sub
            End If
                mfg_clientes.Rows = 2
                mfg_clientes = Clear
                mfg_clientes.FormatString = "Código  |Nome                                                          |Nascimento  |Tel_residêncial  |Celular             |Email                              |Logradouro                                     |Número        |Complemento           "
            tabcli.MoveFirst
            Do While tabcli.EOF = False
                c_e_p = tabcli!Cep
                mfg_clientes.TextMatrix(mfg_clientes.Rows - 1, 0) = tabcli!codigo
                mfg_clientes.TextMatrix(mfg_clientes.Rows - 1, 1) = tabcli!nome
                mfg_clientes.TextMatrix(mfg_clientes.Rows - 1, 2) = tabcli!Nascimento
                mfg_clientes.TextMatrix(mfg_clientes.Rows - 1, 3) = Format(tabcli!Tel_res, "(&&)&&&&-&&&&")
                mfg_clientes.TextMatrix(mfg_clientes.Rows - 1, 4) = Format(tabcli!Celular, "(&&)&&&&-&&&&")
                mfg_clientes.TextMatrix(mfg_clientes.Rows - 1, 5) = tabcli!Email
                Call converter
                mfg_clientes.TextMatrix(mfg_clientes.Rows - 1, 6) = L
                mfg_clientes.TextMatrix(mfg_clientes.Rows - 1, 7) = tabcli!Numero
                mfg_clientes.TextMatrix(mfg_clientes.Rows - 1, 8) = tabcli!Complemento
                tabcli.MoveNext
                mfg_clientes.Rows = mfg_clientes.Rows + 1
                Loop
                mfg_clientes.Rows = mfg_clientes.Rows - 1
            'Else
             '   mfg_clientes.Rows = 2
              '  mfg_clientes = Clear
               ' mfg_clientesFormatString = "Código  |Nome                                                          |Nascimento  |Tel_residêncial  |Celular             |Email                              |Logradouro                                     |Número        |Complemento           "
            'End If
End Sub
Private Sub converter()
            Call abrir
            tabcli2.Close
            tabcli2.Open "select * from Clientes where Cep like '" & c_e_p & "'"
            If tabcli2.RecordCount > 0 Then
            tablocalizacao.Close
            tablocalizacao.Open "Select * from  Localizacoes where Cep = '" & c_e_p & "'"
            L = tablocalizacao!logradouro
            End If
End Sub
Private Sub msk_cep_LostFocus()
            Call desabilitar_mascara
            If tablocalizacao.State = adStateOpen Then tablocalizacao.Close
            tablocalizacao.Open "select * from Localizacoes where Cep = '" & msk_cep & "'"
            If tablocalizacao.RecordCount = 1 Then
            txt_logradouro.Text = tablocalizacao!logradouro
            cod_b = tablocalizacao!Cod_bairro
            End If
            If tablocalizacao.RecordCount = 0 Then
            If MsgBox("Esta localização ainda não está cadastrado! Deseja cadastrar agora?", vbYesNo + vbDefaultButton1 + vbQuestion, "Arbimy manager 2.0") = vbYes Then frm_localizacoes.Show
            Exit Sub
            txt_bairro = Clear
            txt_cidade = Clear
            txt_logradouro = Clear
            txt_estado = Clear
            msk_cep.Text = Empty
            Exit Sub
            Else
            End If
            If tabbairro.State = adStateOpen Then tabbairro.Close
            tabbairro.Open "select * from Bairros where Cod_bairro = " & cod_b
            If tabbairro.RecordCount = 1 Then
            cod_c = tabbairro!Cod_cidade
            txt_bairro.Text = tabbairro!bairro
            End If
            If tabcid.State = adStateOpen Then tabcid.Close
            tabcid.Open "select * from Cidades where Cod_cidade = " & cod_c
            If tabcid.RecordCount = 1 Then
            cod_u = tabcid!Cod_estado
            txt_cidade.Text = tabcid!cidade
            End If
            If tabuf.State = adStateOpen Then tabuf.Close
            tabuf.Open "select * from Ufs where Codigo = " & cod_u
            If tabuf.RecordCount = 1 Then
            txt_uf.Text = tabuf!estado
              Call habilitar_mascara
            End If
End Sub
Private Sub msk_rg_LostFocus()
            cmd_salvar.Enabled = True
            msk_rg = UCase(msk_rg)
End Sub
Private Sub txt_nome_LostFocus()
            txt_nome = UCase(txt_nome)
End Sub
