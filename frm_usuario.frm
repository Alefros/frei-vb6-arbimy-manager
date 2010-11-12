VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_usuario 
   Caption         =   "Cadastro de usúarios"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   6330
   Icon            =   "frm_usuario.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   6330
   Begin MSFlexGridLib.MSFlexGrid mfg_usuarios 
      Height          =   1815
      Left            =   0
      TabIndex        =   26
      Top             =   3240
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   3201
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Caption         =   "Comandos"
      Height          =   3255
      Left            =   5160
      TabIndex        =   19
      Top             =   0
      Width           =   1095
      Begin VB.CommandButton cmd_excluir 
         Height          =   615
         Left            =   240
         Picture         =   "frm_usuario.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Excluir"
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton cmd_alterar 
         Height          =   615
         Left            =   240
         Picture         =   "frm_usuario.frx":2AAC
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Alterar"
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton comd_salvar 
         Height          =   615
         Left            =   240
         Picture         =   "frm_usuario.frx":2DB6
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Salvar"
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmd_novo 
         Height          =   615
         Left            =   240
         Picture         =   "frm_usuario.frx":30C0
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Novo"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Informações de acesso"
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
      Left            =   0
      TabIndex        =   12
      Top             =   2040
      Width           =   5055
      Begin VB.TextBox txt_login 
         Height          =   375
         Left            =   1080
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cbo_funcao 
         Height          =   315
         Left            =   3480
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin MSMask.MaskEdBox msk_senha 
         Height          =   375
         Left            =   1080
         TabIndex        =   17
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   6
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_codigo 
         Height          =   375
         Left            =   3480
         TabIndex        =   21
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   6
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
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
         Left            =   2640
         TabIndex        =   20
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Função"
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
         Left            =   2640
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Senha"
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
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Login"
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
         Top             =   360
         Width           =   735
      End
   End
   Begin MSMask.MaskEdBox msk_telefone 
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   13
      Mask            =   "(99)9999-9999"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contatos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   -120
      TabIndex        =   7
      Top             =   1200
      Width           =   5175
      Begin MSMask.MaskEdBox msk_celular 
         Height          =   375
         Left            =   3600
         TabIndex        =   10
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "(99)9999-9999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
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
         Left            =   2760
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label4 
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
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.TextBox txt_rg 
      Height          =   375
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txt_nome 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   360
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informações pessoais do usúario"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin MSMask.MaskEdBox msk_cpf 
         Height          =   375
         Left            =   3360
         TabIndex        =   6
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   14
         Mask            =   "999.999.999-99"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Cpf"
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
         Left            =   2640
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Rg"
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
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
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
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frm_usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
