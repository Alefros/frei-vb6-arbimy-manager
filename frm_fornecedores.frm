VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_fornecedores 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Fornecedores"
   ClientHeight    =   8340
   ClientLeft      =   -15
   ClientTop       =   360
   ClientWidth     =   12885
   Icon            =   "frm_fornecedores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid mfg_fornecedores 
      Height          =   4215
      Left            =   0
      TabIndex        =   19
      Top             =   3960
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   7435
      _Version        =   393216
      BackColorFixed  =   14737632
      BackColorBkg    =   12632256
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informações bancárias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   46
      Top             =   2400
      Width           =   5775
      Begin VB.ComboBox cbo_banco 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txt_conta 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txt_agencia 
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Conta corrente"
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
         TabIndex        =   49
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label24 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Agência"
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
         TabIndex        =   48
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Banco"
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
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Comandos"
      Height          =   2175
      Left            =   10800
      TabIndex        =   44
      Top             =   240
      Width           =   2055
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
         Picture         =   "frm_fornecedores.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Novo"
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
         Left            =   120
         Picture         =   "frm_fornecedores.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Salvar"
         Top             =   1200
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
         Left            =   1200
         Picture         =   "frm_fornecedores.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Alterar"
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
         Left            =   1200
         Picture         =   "frm_fornecedores.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Excluir"
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Contato do fornecedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5880
      TabIndex        =   38
      Top             =   2400
      Width           =   6975
      Begin VB.TextBox txt_email 
         Height          =   315
         Left            =   1560
         TabIndex        =   15
         Top             =   1080
         Width           =   5295
      End
      Begin VB.TextBox txt_falarcom2 
         Height          =   285
         Left            =   4920
         TabIndex        =   14
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txt_falarcom1 
         Height          =   285
         Left            =   4920
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin MSMask.MaskEdBox msk_telcom 
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "(99)9999-9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_telcom2 
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "(99)9999-9999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label18 
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
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Falar com 2"
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
         Left            =   3600
         TabIndex        =   42
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Falar com"
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
         Left            =   3600
         TabIndex        =   41
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tel com 2"
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
         TabIndex        =   40
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tel com"
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
         TabIndex        =   39
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Endereço do fornecedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      TabIndex        =   26
      Top             =   240
      Width           =   5775
      Begin VB.TextBox txt_numero 
         Height          =   405
         Left            =   4680
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txt_complemento 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txt_cidade 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   34
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox txt_bairro 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   32
         Top             =   1440
         Width           =   3975
      End
      Begin VB.TextBox txt_logradouro 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   30
         Top             =   1800
         Width           =   3975
      End
      Begin VB.TextBox txt_uf 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         TabIndex        =   28
         Top             =   720
         Width           =   975
      End
      Begin MSMask.MaskEdBox msk_cep 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "99999-999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Complemento"
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
         TabIndex        =   37
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Numero"
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
         Left            =   3720
         TabIndex        =   36
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label11 
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
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label10 
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
         Left            =   120
         TabIndex        =   33
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label9 
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
         Left            =   120
         TabIndex        =   31
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label8 
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
         Left            =   3720
         TabIndex        =   29
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label7 
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
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.TextBox txt_codigo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   22
      Top             =   0
      Width           =   1935
   End
   Begin VB.Frame Frame1 
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
      Height          =   2175
      Left            =   5880
      TabIndex        =   20
      Top             =   240
      Width           =   4815
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txt_razaosocial 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   3135
      End
      Begin MSMask.MaskEdBox msk_ie 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   960
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   15
         Mask            =   "999.999.999.999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_cnpj 
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   1440
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   18
         Mask            =   "99.999.999/9999-99"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
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
         TabIndex        =   25
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label4 
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
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
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
         TabIndex        =   23
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         TabIndex        =   21
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      Caption         =   "(*) Preenchimento obrigatório"
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
      Left            =   5880
      TabIndex        =   50
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label2 
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
      TabIndex        =   45
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "frm_fornecedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L_Colunas, L_linha As Long
Dim codfor As Integer
Dim L_Codfor
Public tabfor As New ADODB.Recordset
Public cod_for2 As New ADODB.Recordset
Private Sub desabilitar_mascara()
        msk_ie.PromptInclude = False
        msk_cnpj.PromptInclude = False
        msk_cep.PromptInclude = False
        msk_telcom.PromptInclude = False
        msk_telcom2.PromptInclude = False
End Sub
Private Sub habilitar_mascara()
        msk_ie.PromptInclude = True
        msk_cnpj.PromptInclude = True
        msk_cep.PromptInclude = True
        msk_telcom.PromptInclude = True
        msk_telcom2.PromptInclude = True
End Sub
Private Sub cmd_alterar_Click()
             status = "alteradas"
            If tabfor.State = adStateOpen Then tabcli.Close
                tabcli.Open "Select * From Clientes where Codigo like '" & txt_codigo & "'"
                If tabcli.RecordCount <> 0 Then
                'Call alterar
                End If
End Sub

Private Sub cmd_novo_Click()
            Call cod_for
            Call limpar_fornecedor
            msk_cep.SetFocus
End Sub
Private Sub cod_for()
            
            codfor = 1
A:
            If cod_for2.State = adStateOpen Then cod_for2.Close
            cod_for2.Open "Fornecedores", conectar, adOpenKeyset, adLockOptimistic
            If cod_for2.State = adStateOpen Then cod_for2.Close
            cod_for2.Open "select * from Fornecedores where Codigo = " & codfor
            If cod_for2.RecordCount > 0 Then
                codfor = codfor + 1
                GoTo A
            End If
            txt_codigo = codfor
End Sub
Private Sub cmd_salvar_Click()
             status = "gravadas"
            If txt_razaosocial = Empty Or msk_ie = Empty Or msk_cep = Empty Or txt_numero = Empty Or txt_email = Empty Or msk_telcom = Empty Or msk_cnpj = Empty Then
                MsgBox "Descupe-nos! Há informações que não foram preenchidas", vbExclamation, "Tricon supermercados LTDA."
                Exit Sub
            End If
                Call gravar
                Call box
                Call limpar_fornecedor
                Call carregar_lista
End Sub

Private Sub Form_Load()
            Call carregar_lista
            If cod_for2.State = adStateOpen Then cod_for2.Close
            cod_for2.Open "Fornecedores", conectar, adOpenKeyset, adLockOptimistic
            Call cod_for
End Sub
Private Sub limpar_fornecedor()
            Call desabilitar_mascara
              msk_cep = Clear
              msk_ie = Clear
              msk_cnpj = Clear
              msk_telcom = Clear
              msk_telcom2 = Clear
            txt_codigo = Clear
            txt_numero = Clear
            txt_complemento = Clear
            txt_uf = Clear
            txt_cidade = Clear
            txt_bairro = Clear
            txt_logradouro = Clear
            txt_conta = Clear
            txt_agencia = Clear
            txt_nome = Clear
            txt_razaosocial = Clear
            txt_falarcom = Clear
            txt_falarcom2 = Clear
            txt_email = Clear
              cbo_banco = Clear
            Call habilitar_mascara
            Call cod_for
End Sub

Private Sub Form_Unload(Cancel As Integer)
            Call voltar_botao
End Sub
Private Sub gravar()
            Call abrir
            Call desabilitar_mascara
            If tabfor.State = adStateOpen Then tabfor.Close
            tabfor.Open "select * from Fornecedores where Ie like " & msk_ie.Text
            If tabfor.RecordCount = 0 Then
            If status <> "alteradas" Then cod_for2.AddNew
            cod_for2!Cnpj = msk_cnpj.Text
            cod_for2!Razao_social = txt_razaosocial
            cod_for2!Numero = txt_numero
            cod_for2!Complemento = txt_complemento
            cod_for2!Email = txt_email
            cod_for2!Tel_com1 = msk_telcom.Text
            cod_for2!Tel_com2 = msk_telcom2.Text
            cod_for2!Falar_com = txt_falarcom1
            cod_for2!Falar_com_2 = txt_falarcom2
            cod_for2!Ie = msk_ie.Text
            cod_for2!codigo = txt_codigo
            cod_for2!Cep = msk_cep.Text
            cod_for2!Conta_corrente = txt_conta
            cod_for2!nome = txt_nome
            cod_for2.Update
            Call cod_for
            Call habilitar_mascara
            End If
End Sub
Private Sub carregar_lista()
            Call abrir
            If tabfor.State = adStateOpen Then tabfor.Close
            tabfor.Open "select * from Fornecedores"
            If tabfor.RecordCount = 0 Then
                MsgBox "Descupe-nos, não temos nenhum fornecedor cadastrado", vbInformation, "Arbimy manager 2.0"
                Exit Sub
            End If
                mfg_fornecedores.Rows = 2
                mfg_fornecedores = Clear
                mfg_fornecedores.FormatString = "Código  |Fornecedor                                                          |CNPJ  |Inscrição estadual  |Telefone comercial             |Email                              |Logradouro                                     |Número        |Complemento           "
            tabfor.MoveFirst
            Do While tabfor.EOF = False
                
                mfg_fornecedores.TextMatrix(mfg_fornecedores.Rows - 1, 0) = tabfor!codigo
                mfg_fornecedores.TextMatrix(mfg_fornecedores.Rows - 1, 1) = tabfor!fornecedor
                mfg_fornecedores.TextMatrix(mfg_fornecedores.Rows - 1, 2) = tabfor!Cnpj
                mfg_fornecedores.TextMatrix(mfg_fornecedores.Rows - 1, 3) = tabfor!Ie
                mfg_fornecedores.TextMatrix(mfg_fornecedores.Rows - 1, 4) = tabfor!Tel_com1
                mfg_fornecedores.TextMatrix(mfg_fornecedores.Rows - 1, 5) = tabfor!Email
                'mfg_fornecedores.TextMatrix(mfg_clientes.Rows - 1, 6) = L
                mfg_fornecedores.TextMatrix(mfg_fornecedores.Rows - 1, 7) = tabfor!Numero
                mfg_fornecedores.TextMatrix(mfg_fornecedores.Rows - 1, 8) = tabfor!Complemento
                tabfor.MoveNext
                mfg_fornecedores.Rows = mfg_fornecedores.Rows + 1
                Loop
                mfg_fornecedores.Rows = mfg_fornecedores.Rows - 1
End Sub
Private Sub abrir()
            If tabfor.State = adStateOpen Then tabfor.Close
            tabfor.Open "Fornecedores", conectar, adOpenKeyset, adLockOptimistic
            If cod_for2.State = adStateOpen Then cod_for2.Close
            cod_for2.Open "Fornecedores", conectar, adOpenKeyset, adLockOptimistic
            If tabuf.State = adStateOpen Then tabuf.Close
                tabuf.Open "Ufs", conectar, adOpenKeyset, adLockOptimistic
            If tabcid.State = adStateOpen Then tabcid.Close
                tabcid.Open "Cidades", conectar, adOpenKeyset, adLockOptimistic
            If tabbairro.State = adStateOpen Then tabbairro.Close
                tabbairro.Open "Bairros", conectar, adOpenKeyset, adLockOptimistic
            If tablocalizacao.State = adStateOpen Then tablocalizacao.Close
                tablocalizacao.Open "Localizacoes", conectar, adOpenKeyset, adLockOptimistic


End Sub

Private Sub mfg_fornecedores_Click()
            L_linha = mfg_fornecedores.Row
            L_Codfor = mfg_fornecedores.TextMatrix(L_linha, 0)
            Call abrir
            tabfor.Close
            tabfor.Open "Select * From  Fornecedores Where Codigo = " & L_Codfor
            Call Mostrar
            cmd_alterar.Enabled = True
            cmd_excluir.Enabled = True
End Sub
Private Sub Mostrar()
            Call abrir
            desabilitar_mascara
            msk_ie.Text = tabfor!Ie
            msk_cnpj.Text = tabfor!Cnpj
            msk_cep.Text = tabfor!Cep
            msk_telcom.Text = tabfor!Tel_com1
            msk_telcom2.Text = tabfor!Tel_com2
            txt_codigo.Text = tabfor!codigo
            txt_numero = tabfor!Numero
            txt_complemento = tabfor!Complemento
            'txt_uf
            'txt_cidade
            'txt_bairro
            'txt_logradouro
            'txt_conta
            'txt_agencia
            'txt_nome
            txt_razaosocial = tabfor!Razao_social
            'txt_falarcom
            'txt_falarcom2
            'txt_email
            'cbo_banco
            habilitar_mascara
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

