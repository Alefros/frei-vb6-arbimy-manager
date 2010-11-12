VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_segmentos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de segmentos"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4350
   FontTransparent =   0   'False
   HasDC           =   0   'False
   Icon            =   "frm_segmentos.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid mfg_segmentos 
      Height          =   2055
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3625
      _Version        =   393216
      FormatString    =   "              Código|                                               Segmento  "
   End
   Begin VB.TextBox txt_segmento 
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   480
      Width           =   3135
   End
   Begin VB.Frame Frame5 
      Caption         =   "Comandos"
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   3120
      Width           =   4335
      Begin VB.CommandButton cmd_novo 
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
         Picture         =   "frm_segmentos.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Novo"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_salvar 
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
         Left            =   1320
         Picture         =   "frm_segmentos.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salvar"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_alterar 
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
         Left            =   2280
         Picture         =   "frm_segmentos.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Alterar"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_excluir 
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
         Left            =   3240
         Picture         =   "frm_segmentos.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Excluir"
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox txt_codigo 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label2 
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
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   1215
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "frm_segmentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L_Colunas, L_Linha As Long
Dim L_Codsegmento As Integer
Public Codsegmento As New ADODB.Recordset
Private Sub cmd_alterar_Click()
            status = "alteradas"
            If tabsegmento.State = adStateOpen Then tabsegmento.Close
                tabsegmento.Open "Select * From Segmentos where Codigo like '" & txt_codigo & "'"
                If tabsegmento.RecordCount <> 0 Then
                Call alterar
                End If
            Call box
            Call carregar_lista
            cmd_alterar.Enabled = False
            cmd_salvar.Enabled = False
            cmd_excluir.Enabled = False
End Sub
Private Sub cmd_excluir_Click()
            status = "excluidas"
            If MsgBox("Deseja realmente excluir estas informações?", vbYesNo + vbDefaultButton2 + vbQuestion, "Arbimy manager 2.0") = vbYes Then
                If tabsegmento.State = adStateOpen Then tabsegmento.Close
                    tabsegmento.Open "select * from Segmentos where Codigo = " & txt_codigo
                    If tabsegmento.RecordCount <> 0 Then
                        conectar.Execute "Delete From Segmentos where Codigo like '" & txt_codigo & "'"
              Call box
              Call carregar_lista
              Call limpar
              cmd_alterar.Enabled = False
              cmd_excluir.Enabled = False
                End If
                    End If
End Sub
Private Sub cmd_novo_Click()
            Call limpar
            cmd_salvar.Enabled = False
            cmd_alterar.Enabled = False
            cmd_excluir.Enabled = False
            
End Sub
Private Sub cmd_salvar_Click()
            status = "gravadas"
            If txt_segmento = Empty Then
                MsgBox "Descupe-nos! Há informações que não foram preenchidas", vbExclamation, "Arbimy manager 2.0"
                Exit Sub
            End If
            Call gravar
            Call box
            Call limpar
            Call carregar_lista
End Sub
Private Sub Form_Load()
            Call abrir
            Call cod_segmento
            Call carregar_lista
End Sub

Private Sub Form_Unload(Cancel As Integer)
            Call voltar_botao
End Sub

Private Sub mfg_segmentos_Click()
            L_Linha = mfg_segmentos.Row
            L_Codsegmento = mfg_segmentos.TextMatrix(L_Linha, 0)
            Call abrir
            tabsegmento.Close
            tabsegmento.Open "select * from Segmentos where Codigo = " & L_Codsegmento
            Call Mostrar
End Sub
Private Sub txt_segmento_LostFocus()
            cmd_salvar.Enabled = True
End Sub
Private Sub limpar()
            txt_segmento = Clear
            txt_segmento.SetFocus
            Call cod_segmento
End Sub
Private Sub gravar()
            Call abrir
            tabsegmento.Close
            tabsegmento.Open "select * from Segmentos where Codigo = " & txt_codigo
            If tabsegmento.RecordCount = 0 Then
            If status <> "alteradas" Then Codsegmento.AddNew
            Codsegmento!Codigo = txt_codigo.Text
            Codsegmento!Segmento = txt_segmento.Text
            Codsegmento.Update
            End If
End Sub
Private Sub abrir()
            If tabsegmento.State = adStateOpen Then tabsegmento.Close
            tabsegmento.Open "Segmentos", conectar, adOpenKeyset, adLockOptimistic
            If Codsegmento.State = adStateOpen Then Codsegmento.Close
            Codsegmento.Open "Segmentos", conectar, adOpenKeyset, adLockOptimistic
End Sub
Private Sub cod_segmento()
            Dim cod_segmento As Integer
            cod_segmento = 1
a:
            Call abrir
            If Codsegmento.State = adStateOpen Then Codsegmento.Close
            Codsegmento.Open "select * from Segmentos where Codigo = " & cod_segmento
            If Codsegmento.RecordCount > 0 Then
                cod_segmento = cod_segmento + 1
                GoTo a
            End If
            txt_codigo = cod_segmento
End Sub
Private Sub carregar_lista()
            Call abrir
            If tabsegmento.State = adStateOpen Then tabsegmento.Close
            tabsegmento.Open "select * from Segmentos"
            If tabsegmento.RecordCount = 0 Then
                MsgBox "Não há segmentos cadastrados", vbInformation, "Arbimy manager 2.0"
                Exit Sub
            End If
                mfg_segmentos.Rows = 2
                mfg_segmentos = Clear
                mfg_segmentos.FormatString = "Código  |Segmento                      "
            tabsegmento.MoveFirst
            Do While tabsegmento.EOF = False
                mfg_segmentos.TextMatrix(mfg_segmentos.Rows - 1, 0) = tabsegmento!Codigo
                mfg_segmentos.TextMatrix(mfg_segmentos.Rows - 1, 1) = tabsegmento!Segmento
                tabsegmento.MoveNext
                mfg_segmentos.Rows = mfg_segmentos.Rows + 1
                Loop
                mfg_segmentos.Rows = mfg_segmentos.Rows - 1
End Sub
Private Sub Mostrar()
            txt_codigo.Text = tabsegmento!Codigo
            txt_segmento.Text = tabsegmento!Segmento
            cmd_alterar.Enabled = True
            cmd_excluir.Enabled = True
End Sub
Private Sub alterar()
            tabsegmento!Segmento = txt_segmento.Text
            tabsegmento.Update
            Call limpar
End Sub


