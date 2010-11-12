VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_funcionarios 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de funcionários"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6180
   Icon            =   "frm_usuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Endereço"
      Height          =   1335
      Left            =   0
      TabIndex        =   36
      Top             =   1200
      Width           =   6135
      Begin VB.TextBox txt_uf 
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txt_logradouro 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   13
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txt_bairro 
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   720
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txt_cidade 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox txt_complemento 
         Height          =   285
         Left            =   4680
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txt_numero 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_cep 
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "99999-999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cep"
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
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Caption         =   "UF"
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
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Logradouro"
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
         Left            =   2400
         TabIndex        =   41
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bairro"
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
         Left            =   120
         TabIndex        =   40
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cidade"
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
         Left            =   1560
         TabIndex        =   39
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nº"
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
         Left            =   1920
         TabIndex        =   38
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label13 
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
         Left            =   3480
         TabIndex        =   37
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ComboBox cbo_funcao 
      Height          =   315
      Left            =   4920
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid mfg_funcionarios 
      Height          =   2295
      Left            =   0
      TabIndex        =   19
      Top             =   3840
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4048
      _Version        =   393216
      FormatString    =   ""
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Comandos"
      Height          =   1095
      Left            =   0
      TabIndex        =   25
      Top             =   6120
      Width           =   6135
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
         Left            =   720
         Picture         =   "frm_usuario.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   2040
         Picture         =   "frm_usuario.frx":2AAC
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Salvar"
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
         Left            =   3360
         Picture         =   "frm_usuario.frx":2DB6
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Left            =   4680
         Picture         =   "frm_usuario.frx":30C0
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Excluir"
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informações de acesso"
      Height          =   1335
      Left            =   0
      TabIndex        =   24
      Top             =   2520
      Width           =   6135
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frm_usuario.frx":33CA
         Left            =   5040
         List            =   "frm_usuario.frx":33CC
         TabIndex        =   46
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txt_confirmar 
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4560
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   17
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox cbo_acesso 
         Height          =   315
         ItemData        =   "frm_usuario.frx":33CE
         Left            =   5040
         List            =   "frm_usuario.frx":33D8
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txt_senha 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   16
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txt_login 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   15
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Permissão"
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
         Left            =   3000
         TabIndex        =   47
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirmar senha"
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
         Left            =   3000
         TabIndex        =   45
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Deseja criar conta de acesso ao sistema para este funcionário? "
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Nova senha"
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
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
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
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Contatos"
      Height          =   975
      Left            =   3840
      TabIndex        =   18
      Top             =   240
      Width           =   2295
      Begin MSMask.MaskEdBox msk_celular 
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "(99)9999-9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_telefone 
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "(99)9999-9999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Celular"
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
         TabIndex        =   30
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone"
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
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informações pessoais do funcionário"
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3735
      Begin VB.TextBox txt_nome 
         Height          =   225
         Left            =   720
         MaxLength       =   50
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
      Begin MSMask.MaskEdBox msk_cpf 
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   14
         Mask            =   "999.999.999-99"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_rg 
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99.999.999-9"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cpf"
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
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rg"
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
         Left            =   2160
         TabIndex        =   27
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
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
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
   End
   Begin MSMask.MaskEdBox msk_codigo 
      Height          =   255
      Left            =   720
      TabIndex        =   35
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   6
      PromptChar      =   "_"
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Função"
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
      Left            =   3840
      TabIndex        =   34
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
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
      Left            =   120
      TabIndex        =   33
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frm_funcionarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tabfuncionario As New ADODB.Recordset
Public tabusuario As New ADODB.Recordset
Public tabfuncao As New ADODB.Recordset
Dim codusuario, funcao2 As Integer
Dim L_codusu, codfuncao As String
Dim L_linha As String
Private Sub codusu()
            codusuario = 1
A:
            If tabfuncionario.State = adStateOpen Then tabfuncionario.Close
            tabfuncionario.Open "select * from Funcionarios where Codigo = " & codusuario
            If tabfuncionario.RecordCount > 0 Then
                codusuario = codusuario + 1
                GoTo A:
            End If
            msk_codigo = codusuario
End Sub

Private Sub cbo_acesso_Click()
            If cbo_acesso.Text = "Sim" Then
            txt_login.Enabled = True
            txt_senha.Enabled = True
            txt_confirmar.Enabled = True
            ElseIf cbo_acesso.Text = "Não" Then
            txt_login.Enabled = False
            txt_senha.Enabled = False
            txt_confirmar.Enabled = False
            End If
End Sub

Private Sub cmd_alterar_Click()
            status = "alteradas"
            Call gravar_funcao
            Call box
            'Call apagar
            Call codusu
End Sub
Private Sub cmd_excluir_Click()

status = "excluídas"
            If tabusuario.State = adStateOpen Then tabusuario.Close
            tabusuario.Open "funcoes", conectar, adOpenKeyset, adLockOptimistic
                If tabusuario.State = adStateOpen Then tabusuario.Close
                tabusuario.Open "select * from Usuarios where Codigo = " & msk_codigo
                If tabusuario.RecordCount = 0 Then MsgBox "DIGITE UM USUÁRIO EXISTENTE": txt_usuario.SetFocus: Exit Sub
            If tabusuario.State = adStateOpen Then tabusuario.Close
            tabusuario.Open "Usuarios", conectar, adOpenKeyset, adLockOptimistic
            L_Codfunc = txt_codigo
            If tabusuario.State = adStateOpen Then tabusuario.Close
            If MsgBox("Deseja realmente excluir este Usuário ?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
               tabusuario.Open "Select * From Usuarios Where Codigo = " & msk_codigo
               If tabusuario.RecordCount = 1 Then
                  conectar.Execute "Delete From Usuarios Where Codigo = " & msk_codigo
                  Call flex
               End If
               Call box
            End If
           ' Call apagar
            Call codusu
            
End Sub

Private Sub cmd_novo_Click()
            Call desativar
            Call limpar
            If cbo_funcao.ListCount > 0 Then cbo_funcao.ListIndex = 0
            Call carregar_combo
            Call flex
            Call ativar
            Call codusu
End Sub

Private Sub cmd_salvar_Click()
            If txt_senha <> txt_confirmar Then
                MsgBox "As senhas são incompatíveis!", vbExclamation, "Arbimy manager 2.0"
                txt_senha.BackColor = &HFF&
                txt_confirmar.BackColor = &HFF&
                txt_senha = Clear
                txt_confirmar = Clear
                Exit Sub
            Else
            status = "gravadas"
            Call gravar_funcao
            Call box
            End If
End Sub
Private Sub Form_Load()
            Call abrir_banco
            Call abrir
            Call codusu
            Call carregar_combo
            Call flex
           
End Sub

Private Sub desativar()
            msk_celular.PromptInclude = False
            msk_codigo.PromptInclude = False
            msk_cpf.PromptInclude = False
            msk_rg.PromptInclude = False
            msk_telefone.PromptInclude = False
            msk_cep.PromptInclude = False
End Sub

Private Sub ativar()
        msk_celular.PromptInclude = True
        msk_codigo.PromptInclude = True
        msk_cpf.PromptInclude = True
        msk_rg.PromptInclude = True
        msk_telefone.PromptInclude = True
        msk_cep.PromptInclude = True
End Sub
Private Sub gravar_funcao()
            Call desativar
                If status = "alteradas" Then
            If tabfuncionario.State = adStateOpen Then tabfuncionario.Close: tabfuncionario.Open "select * from Funcionarios where Codigo = " & msk_codigo
                Else
                tabfuncionario.AddNew
                End If
                    tabfuncionario!nome = txt_nome
                    tabfuncionario!Rg = msk_rg
                    tabfuncionario!CPF = msk_cpf
                    tabfuncionario!Cep = msk_cep
                    tabfuncionario!Telefone = msk_telefone
                    tabfuncionario!Celular = msk_celular
                    tabfuncionario!Numero = txt_numero
                        codfuncao = cbo_funcao.Text
                        If tabfuncao.State = 1 Then tabfuncao.Close
                        tabfuncao.Open "Select * from funcoes where Funcao like '" & codfuncao & "'"
                        If tabfuncao.RecordCount = 1 Then
                            funcao2 = tabfuncao!codigo
                        End If
                    tabfuncionario!Funcao = funcao2
                    tabfuncionario!Complemento = txt_complemento
                    tabfuncionario!codigo = msk_codigo
                    tabfuncionario.Update
            
            If tabusuario.State = adStateOpen Then tabusuario.Close: tabusuario.Open "select * from Usuarios where Cod_funcionario = " & msk_codigo
            'Else
                tabusuario.AddNew
                'End If
                tabusuario!Login = txt_login
                tabusuario!Senha = txt_senha
                tabusuario!Cod_funcionario = msk_codigo
                tabusuario.Update
            
            Call ativar
            Call flex
End Sub
Private Sub abrir_conexao()
            If tabfuncionario.State = adStateOpen Then tabfuncionario.Close
            tabfuncionario.Open "Funcionarios", conectar, adOpenKeyset, adLockOptimistic

            If tabfuncao.State = adStateOpen Then tabfuncao.Close
            tabfuncao.Open "Funcoes", conectar, adOpenKeyset, adLockOptimistic
            
            If tabusuario.State = adStateOpen Then tabusuario.Close
            tabusuario.Open "Usuarios", conectar, adOpenKeyset, adLockOptimistic

End Sub
Private Sub fechar_conexao()
            If tabusuario.State = adStateOpen Then tabusuario.Close
            If tabfuncao.State = adStateOpen Then tabfuncao.Close
End Sub
Private Sub carregar_combo()
                Call abrir_conexao
                cbo_funcao.Clear
                cbo_funcao.Text = tabfuncao!Funcao
                Do Until tabfuncao.EOF = True
                    cbo_funcao.AddItem tabfuncao!Funcao
                    cbo_funcao.ItemData(cbo_funcao.NewIndex) = tabfuncao!codigo
                    tabfuncao.MoveNext
                Loop
End Sub
Private Sub flex()
           Call abrir
           If tabfuncionario.State = 1 Then tabfuncionario.Close
           tabfuncionario.Open "select * from Funcionarios"
           If tabfuncionario.RecordCount = 0 Then
                MsgBox "Nenhum funcionário cadastrado", vbInformation, "Arbimy manager 2.0"
                Exit Sub
            End If
                mfg_funcionarios.Clear
                mfg_funcionarios.Rows = 2
                mfg_funcionarios.FormatString = "Código|  Nome                                     |Função        |Rg                      "
            If tabfuncionario.State = adStateOpen Then tabfuncionario.Close
                tabfuncionario.Open "select * from Funcionarios order by Codigo"
            Do While tabfuncionario.EOF = False
                mfg_funcionarios.TextMatrix(mfg_funcionarios.Rows - 1, 0) = tabfuncionario!codigo
                mfg_funcionarios.TextMatrix(mfg_funcionarios.Rows - 1, 1) = tabfuncionario!nome
                    funcao2 = tabfuncionario!Funcao
                    If tabfuncao.State = adStateOpen Then tabfuncao.Close
                    tabfuncao.Open "select * from Funcoes where Codigo like '" & funcao2 & "'"
                    If tabfuncao.RecordCount = 1 Then
                        codfuncao = tabfuncao!Funcao
                    End If
                mfg_funcionarios.TextMatrix(mfg_funcionarios.Rows - 1, 2) = codfuncao
                mfg_funcionarios.TextMatrix(mfg_funcionarios.Rows - 1, 3) = Format(tabfuncionario!Rg, "&&.&&&.&&&-&")
            tabfuncionario.MoveNext
                mfg_funcionarios.Rows = mfg_funcionarios.Rows + 1
            Loop
            mfg_funcionarios.Rows = mfg_funcionarios.Rows - 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
            Call voltar_botao
End Sub

Private Sub mfg_funcionarios_Click()
            If mfg_funcionarios.Rows < 2 Then Exit Sub
            Call desativar
            L_linha = mfg_funcionarios.Row
            L_codusu = mfg_funcionarios.TextMatrix(L_linha, 0)
            If tabfuncionario.State = adStateOpen Then tabfuncionario.Close
            tabfuncionario.Open "Select * From Funcionarios Where Codigo = " & L_codusu
            txt_nome = tabfuncionario!nome
            msk_rg = tabfuncionario!Rg
            msk_cpf = tabfuncionario!CPF
            msk_cep = tabfuncionario!Cep
            msk_telefone = tabfuncionario!Telefone
            msk_celular = tabfuncionario!Celular
            msk_codigo = tabfuncionario!codigo
            txt_numero = tabfuncionario!Numero
            
            If tabusuario.State = adStateOpen Then tabusuario.Close
            tabusuario.Open "Select * From Usuarios Where Cod_funcionario = " & L_codusu
            cbo_acesso.Text = "Sim"
            txt_login = tabusuario!Login
            txt_senha = tabusuario!Senha
            txt_confirmar = tabusuario!Senha
            Call ativar
End Sub
Private Sub limpar()
            Call desativar
            txt_nome = Clear
            msk_rg = Clear
            msk_cpf = Clear
            msk_cep = Clear
            msk_telefone = Clear
            msk_celular = Clear
            txt_login = Clear
            txt_senha = Clear
            msk_codigo = Clear
            txt_nome.SetFocus
            Call ativar
End Sub
Private Sub abrir()
            If tabfuncionario.State = adStateOpen Then tabfuncionario.Close
            tabfuncionario.Open "Funcionarios", conectar, adOpenKeyset, adLockOptimistic
            If tabfuncao.State = adStateOpen Then tabfuncao.Close
            tabfuncao.Open "Funcoes", conectar, adOpenKeyset, adLockOptimistic
            If tabusuario.State = adStateOpen Then tabusuario.Close
            tabusuario.Open "Usuarios", conectar, adOpenKeyset, adLockOptimistic
            
End Sub

Private Sub txt_confirmar_Change()

End Sub
