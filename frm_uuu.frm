VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_usuarios 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Cadastro de usuários"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2175
      Left            =   0
      TabIndex        =   21
      Top             =   2760
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3836
      _Version        =   393216
      BackColorFixed  =   14737632
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Comandos"
      ForeColor       =   &H80000006&
      Height          =   1095
      Left            =   0
      TabIndex        =   16
      Top             =   1680
      Width           =   6015
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
         Picture         =   "frm_uuu.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   3240
         Picture         =   "frm_uuu.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Left            =   4560
         Picture         =   "frm_uuu.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Left            =   720
         Picture         =   "frm_uuu.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informações pessoais"
      Height          =   1335
      Left            =   0
      TabIndex        =   9
      Top             =   360
      Width           =   3135
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txt_codfuncionarios 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin MSMask.MaskEdBox msk_rg 
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "99.999.999-&"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
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
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   12
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cod.Funcionários"
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
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informações de acesso"
      Height          =   1335
      Left            =   3240
      TabIndex        =   1
      Top             =   360
      Width           =   2775
      Begin VB.TextBox txt_confirmar 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txt_senha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txt_login 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "C.Senha"
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
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Senha"
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
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txt_codigo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frm_usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_novo_Click()
            Call limpar
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
            If txt_codfuncionarios = Empty Or txt_login = Empty Then
                MsgBox "Descupe-nos! Há informações que não foram preenchidas", vbExclamation, "Arbimy manager 2.0"
                Exit Sub
            End If
                Call gravar
                Call box
                Call limpar
                Call carregar_lista
End Sub

Private Sub Form_Unload(Cancel As Integer)
            Call voltar_botao
End Sub
Private Sub limpar()
            Call desabilitar_mascara
            txt_codigo = Clear
            txt_codfuncionarios = Clear
            msk_rg.Text = Clear
            txt_nome = Clear
            txt_login = Clear
            txt_senha = Clear
            txt_confirmar = Clear
            txt_codfuncionarios.SetFocus
            Call habilitar_mascara
End Sub
Private Sub habilitar_mascara()
            msk_rg.PromptInclude = True
End Sub
Private Sub desabilitar_mascara()
            msk_rg.PromptInclude = False
End Sub
Private Sub gravar()
                Call desativar
            If status = "alteradas" Then
            If tabusuario.State = adStateOpen Then tabusuario.Close: tabusuario.Open "select * from Usuarios where Cod_funcionario = " & msk_codigo
            Else
                tabusuario.AddNew
            End If
                tabusuario!Login = txt_login
                tabusuario!Senha = txt_senha
                tabusuario!Cod_funcionario = msk_codigo
                tabusuario.Update
End Sub
Private Sub carregar_lista()

End Sub
