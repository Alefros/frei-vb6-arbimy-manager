VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_produtos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de produtos"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      Caption         =   "Comandos"
      Height          =   1935
      Left            =   7560
      TabIndex        =   42
      Top             =   1320
      Width           =   1935
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
         Left            =   1080
         Picture         =   "frm_produtos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Excluir"
         Top             =   1080
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
         Left            =   1080
         Picture         =   "frm_produtos.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   45
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
         Left            =   120
         Picture         =   "frm_produtos.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Salvar"
         Top             =   1080
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
         TabIndex        =   43
         ToolTipText     =   "Novo"
         Top             =   240
         Width           =   735
      End
   End
   Begin MSFlexGridLib.MSFlexGrid mfg_produtos 
      Height          =   2775
      Left            =   0
      TabIndex        =   41
      Top             =   4320
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   4895
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Caption         =   "Informações do fornecedor"
      Height          =   975
      Left            =   120
      TabIndex        =   30
      Top             =   3240
      Width           =   9375
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   1320
         TabIndex        =   39
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txt_razaosocial 
         Height          =   285
         Left            =   5760
         TabIndex        =   31
         Top             =   240
         Width           =   3495
      End
      Begin MSMask.MaskEdBox msk_ie 
         Height          =   255
         Left            =   1320
         TabIndex        =   32
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   15
         Mask            =   "999.999.999.999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_cnpj 
         Height          =   255
         Left            =   5760
         TabIndex        =   33
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   18
         Mask            =   "99.999.999/9999-99"
         PromptChar      =   "_"
      End
      Begin VB.Label Label17 
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
         TabIndex        =   40
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label16 
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
         Left            =   4080
         TabIndex        =   36
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label15 
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
         TabIndex        =   35
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label14 
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
         Left            =   4080
         TabIndex        =   34
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   23
      Top             =   2160
      Width           =   3615
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1320
         TabIndex        =   38
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1320
         TabIndex        =   37
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "Qtde Máx."
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
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Qtde Min."
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
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1935
      Left            =   3840
      TabIndex        =   16
      Top             =   1320
      Width           =   3615
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   2040
         TabIndex        =   29
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2040
         TabIndex        =   28
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2040
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Preço de venda"
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
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label10 
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
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label7 
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
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   3615
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "L"
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   120
         Width           =   495
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Kg"
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   120
         Width           =   975
      End
      Begin VB.CheckBox Check3 
         Caption         =   "G"
         Height          =   375
         Left            =   3120
         TabIndex        =   11
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label6 
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
         TabIndex        =   14
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Unidade"
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
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   9375
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   5760
         TabIndex        =   19
         Top             =   600
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         TabIndex        =   18
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   5760
         TabIndex        =   8
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label9 
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
         Left            =   3840
         TabIndex        =   20
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label8 
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
         TabIndex        =   17
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label4 
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
         Left            =   3840
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
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
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   255
      Left            =   5880
      TabIndex        =   3
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   14
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label2 
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
      Left            =   3960
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label1 
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
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "frm_produtos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Sheet1_GotFocus()

End Sub

Private Sub Form_Unload(Cancel As Integer)
        Call voltar_botao
End Sub
