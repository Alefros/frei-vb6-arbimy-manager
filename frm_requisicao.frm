VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frm_requisicao 
   Caption         =   "Requisição de abastecimento"
   ClientHeight    =   8010
   ClientLeft      =   945
   ClientTop       =   2325
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   5685
   Begin VB.Frame Frame2 
      Caption         =   "Informações do operador de estoque"
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
      Left            =   0
      TabIndex        =   26
      Top             =   6000
      Width           =   5655
      Begin VB.CheckBox Check2 
         Caption         =   "Recebida"
         Height          =   495
         Left            =   4320
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Devolvida"
         Height          =   495
         Left            =   3000
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txt_operador 
         Height          =   285
         Left            =   960
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label10 
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
         TabIndex        =   27
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Comandos"
      Height          =   975
      Left            =   0
      TabIndex        =   21
      Top             =   6960
      Width           =   5655
      Begin VB.CommandButton cmd_excluir 
         Height          =   615
         Left            =   4320
         Picture         =   "frm_requisicao.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Excluir"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmd_alterar 
         Height          =   615
         Left            =   3120
         Picture         =   "frm_requisicao.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Alterar"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton comd_salvar 
         Height          =   615
         Left            =   1920
         Picture         =   "frm_requisicao.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Salvar"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmd_novo 
         Height          =   615
         Left            =   720
         Picture         =   "frm_requisicao.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Novo"
         Top             =   240
         Width           =   615
      End
   End
   Begin MSFlexGridLib.MSFlexGrid mfg_requisicao 
      Height          =   2055
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   6
      FormatString    =   "Código|Nome       |Descrição              |Unitário    |Qtde  |Total  "
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informações do produto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   -120
      TabIndex        =   2
      Top             =   3000
      Width           =   5655
      Begin VB.CommandButton cmd_adicionar 
         Caption         =   "Adicionar à lista"
         Height          =   375
         Left            =   2040
         TabIndex        =   20
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox txt_unitario 
         Height          =   375
         Left            =   2040
         TabIndex        =   19
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox txt_qtde 
         Height          =   375
         Left            =   4560
         TabIndex        =   18
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txt_produto 
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txt_descricao 
         Height          =   285
         Left            =   2040
         TabIndex        =   14
         Top             =   1680
         Width           =   3495
      End
      Begin MSMask.MaskEdBox msk_codproduto 
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_codbarras 
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   840
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   14
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         Caption         =   "Qtde"
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
         TabIndex        =   17
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label8 
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
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label7 
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
         TabIndex        =   13
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label6 
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
         TabIndex        =   12
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Código de barras"
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
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label4 
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
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.TextBox txt_codigo 
      Enabled         =   0   'False
      Height          =   405
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin MSMask.MaskEdBox msk_data 
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   0
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "11/11/1111"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox msk_codigo 
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   6
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      Caption         =   "Código funcionário"
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
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
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
      Left            =   4080
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Código requisição"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frm_requisicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
