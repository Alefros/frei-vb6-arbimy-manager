VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_estoque 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Controle de estoque"
   ClientHeight    =   4005
   ClientLeft      =   -15
   ClientTop       =   255
   ClientWidth     =   12900
   Icon            =   "frm_estoque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   12900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "CONTROLE DE ESTOQUE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   11
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox Text3 
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
         Left            =   1080
         TabIndex        =   9
         Top             =   3480
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
         Height          =   2535
         Left            =   8760
         TabIndex        =   2
         Top             =   840
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4471
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         Appearance      =   0
         FormatString    =   "COD_PROD    |QTDE        |P_UNITARIO|TOTAL    "
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   4471
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         Appearance      =   0
         FormatString    =   "DATA          "
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid5 
         Height          =   2535
         Left            =   4920
         TabIndex        =   3
         Top             =   840
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4471
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         Appearance      =   0
         FormatString    =   "COD_PROD    |QTDE        |P_UNITARIO|TOTAL  "
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid6 
         Height          =   2535
         Left            =   1080
         TabIndex        =   4
         Top             =   840
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4471
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         Appearance      =   0
         FormatString    =   "COD_PROD    |QTDE        |P_UNITARIO|TOTAL  "
      End
      Begin VB.Label Label10 
         Caption         =   "TRICON SUPERMERCADOS LTDA."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         TabIndex        =   12
         Top             =   3480
         Width           =   4575
      End
      Begin VB.Label Label9 
         Caption         =   "TOTAL"
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
         Left            =   3240
         TabIndex        =   10
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "STATUS"
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
         Left            =   120
         TabIndex        =   8
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "SALDO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10320
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "SAÍDA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "ENTRADA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frm_estoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
