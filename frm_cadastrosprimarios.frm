VERSION 5.00
Begin VB.Form frm_cadastrosprimarios 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastros primários"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5205
   Icon            =   "frm_cadastrosprimarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_frete 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmd_funcao 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Funções"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmd_marca 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Marcas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmd_segmento 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Segmentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frm_cadastrosprimarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_frete_Click()
            frm_frete.Show
End Sub

Private Sub cmd_funcao_Click()
            frm_funcao.Show
End Sub

Private Sub cmd_marca_Click()
            frm_marcas.Show
End Sub

Private Sub cmd_segmento_Click()
            frm_segmentos.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
            MDIForm1.Tbl_padrao.Buttons.Item(1).Value = tbrUnpressed
End Sub
