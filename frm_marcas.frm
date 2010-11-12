VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_marcas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de marcas"
   ClientHeight    =   4110
   ClientLeft      =   825
   ClientTop       =   2865
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4110
   Begin VB.TextBox txt_marca 
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   480
      Width           =   3135
   End
   Begin VB.Frame Frame5 
      Caption         =   "Comandos"
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   3000
      Width           =   4095
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
         Left            =   3120
         Picture         =   "frm_marcas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
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
         Left            =   2160
         Picture         =   "frm_marcas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Alterar"
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
         Left            =   1200
         Picture         =   "frm_marcas.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salvar"
         Top             =   240
         Width           =   735
      End
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
         Left            =   240
         Picture         =   "frm_marcas.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Novo"
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox txt_codigo 
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid mfg_marcas 
      Height          =   2055
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3625
      _Version        =   393216
      FormatString    =   "              Código|                                               Marca  "
   End
   Begin VB.Label Label2 
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
Attribute VB_Name = "frm_marcas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L_Colunas, L_Linha As Long
Dim L_Codmarca As Integer
Public Codmarca As New ADODB.Recordset

Private Sub cmd_alterar_Click()
            status = "alteradas"
            
            If tabmarca.State = adStateOpen Then tabmarca.Close
                tabmarca.Open "Select * From Marcas where Codigo like '" & txt_codigo & "'"
                If tabmarca.RecordCount <> 0 Then
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
                If tabmarca.State = adStateOpen Then tabmarca.Close
                    tabmarca.Open "select * from Marcas where Codigo = " & txt_codigo
                    If tabmarca.RecordCount <> 0 Then
                        conectar.Execute "Delete From Marcas where Codigo like '" & txt_codigo & "'"
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
            If txt_marca = Empty Then
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
            Call cod_marca
            Call carregar_lista
End Sub

Private Sub Form_Unload(Cancel As Integer)
            Call voltar_botao
End Sub

Private Sub mfg_marcas_Click()
            L_Linha = mfg_marcas.Row
            L_Codmarca = mfg_marcas.TextMatrix(L_Linha, 0)
            Call abrir
            tabmarca.Close
            tabmarca.Open "select * from Marcas where Codigo = " & L_Codmarca
            Call Mostrar
End Sub

Private Sub txt_marca_LostFocus()
            cmd_salvar.Enabled = True
End Sub
Private Sub limpar()
            txt_codigo = Clear
            txt_marca = Clear
            txt_marca.SetFocus
            cod_marca
End Sub
Private Sub gravar()
            Call abrir
            tabmarca.Close
            tabmarca.Open "select * from Marcas where Codigo = " & txt_codigo
            If tabmarca.RecordCount = 0 Then
            If status <> "alteradas" Then Codmarca.AddNew
            Codmarca!Codigo = txt_codigo.Text
            Codmarca!Marca = txt_marca.Text
            Codmarca.Update
            End If
End Sub
Private Sub abrir()
            If tabmarca.State = adStateOpen Then tabmarca.Close
            tabmarca.Open "Marcas", conectar, adOpenKeyset, adLockOptimistic
            If Codmarca.State = adStateOpen Then Codmarca.Close
            Codmarca.Open "Marcas", conectar, adOpenKeyset, adLockOptimistic
End Sub
Private Sub cod_marca()
            Dim cod_marca As Integer
            cod_marca = 1
A:
            Call abrir
            If Codmarca.State = adStateOpen Then Codmarca.Close
            Codmarca.Open "select * from Marcas where Codigo = " & cod_marca
            If Codmarca.RecordCount > 0 Then
                cod_marca = cod_marca + 1
                GoTo A
            End If
            txt_codigo = cod_marca
End Sub
Private Sub carregar_lista()
            Call abrir
            If tabmarca.State = adStateOpen Then tabmarca.Close
            tabmarca.Open "select * from Marcas"
            If tabmarca.RecordCount = 0 Then
                MsgBox "Não há marcas cadastradas", vbInformation, "Arbimy manager 2.0"
                Exit Sub
            End If
                mfg_marcas.Rows = 2
                mfg_marcas = Clear
                mfg_marcas.FormatString = "Código  |Marca                      "
            tabmarca.MoveFirst
            Do While tabmarca.EOF = False
                mfg_marcas.TextMatrix(mfg_marcas.Rows - 1, 0) = tabmarca!Codigo
                mfg_marcas.TextMatrix(mfg_marcas.Rows - 1, 1) = tabmarca!Marca
                tabmarca.MoveNext
                mfg_marcas.Rows = mfg_marcas.Rows + 1
                Loop
                mfg_marcas.Rows = mfg_marcas.Rows - 1
End Sub
Private Sub Mostrar()
            txt_codigo.Text = tabmarca!Codigo
            txt_marca.Text = tabmarca!Marca
            cmd_alterar.Enabled = True
            cmd_excluir.Enabled = True
End Sub
Private Sub alterar()
            tabmarca!Marca = txt_marca.Text
            tabmarca.Update
            Call limpar
End Sub

