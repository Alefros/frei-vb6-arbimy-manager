VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_funcao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de funções"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid mfg_funcoes 
      Height          =   1815
      Left            =   0
      TabIndex        =   9
      Top             =   1920
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3201
      _Version        =   393216
      BackColorFixed  =   14737632
      BackColorBkg    =   12632256
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Comandos"
      Height          =   1095
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   3975
      Begin VB.CommandButton cmd_excluir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Excluir"
         Enabled         =   0   'False
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
         Left            =   3000
         Picture         =   "frm_funcao.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Excluir"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_alterar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Alterar"
         Enabled         =   0   'False
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
         Picture         =   "frm_funcao.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Alterar"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_salvar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Salvar"
         Enabled         =   0   'False
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
         Picture         =   "frm_funcao.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salvar"
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
         Left            =   120
         Picture         =   "frm_funcao.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Novo"
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox txt_funcao 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox txt_codigo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
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
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frm_funcao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L_Colunas, L_Linha As Long
Dim L_Codfunc As Integer
Public tabfuncao As New ADODB.Recordset
Public Codfuncao As New ADODB.Recordset

Private Sub cmd_alterar_Click()
            status = "alteradas"
            If tabfuncao.State = adStateOpen Then tabfuncao.Close
                tabfuncao.Open "Select * From Funcoes where Codigo like '" & txt_codigo & "'"
                If tabfuncao.RecordCount <> 0 Then
                Call alterar
                End If
            Call box
            Call carregar_lista
End Sub

Private Sub cmd_excluir_Click()
            status = "excluidas"
            If MsgBox("Deseja realmente excluir estas informações?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
                If tabfuncao.State = adStateOpen Then tabfuncao.Close
                    tabfuncao.Open "select * from Funcoes where Codigo = " & txt_codigo
                    If tabfuncao.RecordCount <> 0 Then
                        conectar.Execute "Delete From Funcoes where Codigo like '" & txt_codigo & "'"
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
             txt_funcao.SetFocus
             cmd_salvar.Enabled = False
             cmd_alterar.Enabled = False
             cmd_excluir.Enabled = False
End Sub
Private Sub limpar()
            txt_codigo = Clear
            txt_funcao = Clear
            txt_funcao.SetFocus
            Call cod_funcao
End Sub
Private Sub cod_funcao()
            Dim cod_funcao2 As Integer
            cod_funcao2 = 1
A:
            Call abrir
            If Codfuncao.State = adStateOpen Then Codfuncao.Close
            Codfuncao.Open "select * from Funcoes where Codigo = " & cod_funcao2
            If Codfuncao.RecordCount > 0 Then
                cod_funcao2 = cod_funcao2 + 1
                GoTo A
            End If
            txt_codigo = cod_funcao2
End Sub
Private Sub abrir()
            If Codfuncao.State = adStateOpen Then Codfuncao.Close
            Codfuncao.Open "Funcoes", conectar, adOpenKeyset, adLockOptimistic
            If tabfuncao.State = adStateOpen Then tabfuncao.Close
            tabfuncao.Open "Funcoes", conectar, adOpenKeyset, adLockOptimistic
End Sub
Private Sub cmd_salvar_Click()
            status = "gravadas"
            If txt_funcao = Empty Then
                MsgBox "Descupe-nos! Há informações que não foram preenchidas", vbExclamation, "Arbimy manager 2.0"
                Exit Sub
            End If
                Call gravar
                Call box
                Call limpar
                Call carregar_lista
End Sub

Private Sub Form_Load()
            Call cod_funcao
            Call carregar_lista
End Sub

Private Sub Form_Unload(Cancel As Integer)
            Call voltar_botao
End Sub

Private Sub mfg_funcoes_Click()
            L_Linha = mfg_funcoes.Row
            L_Codfunc = mfg_funcoes.TextMatrix(L_Linha, 0)
            Call abrir
            tabfuncao.Close
            tabfuncao.Open "Select * From Funcoes Where Codigo = " & L_Codfunc
            Call Mostrar
            cmd_alterar.Enabled = True
            cmd_excluir.Enabled = True
End Sub

Private Sub txt_funcao_LostFocus()
            If txt_funcao.Text <> Empty Then
            cmd_salvar.Enabled = True
            End If
End Sub
Private Sub gravar()
            If tabfuncao.State = adStateOpen Then tabfuncao.Close
            tabfuncao.Open "select * from Funcoes where Codigo = " & txt_codigo.Text
            If tabfuncao.RecordCount = 0 Then
            If status <> "alteradas" Then Codfuncao.AddNew
            Codfuncao!Codigo = txt_codigo
            Codfuncao!Funcao = txt_funcao.Text
            Codfuncao.Update
            End If
End Sub
Private Sub carregar_lista()
            Call abrir
            If tabfuncao.State = adStateOpen Then tabfuncao.Close
            tabfuncao.Open "select * from Funcoes"
            If tabfuncao.RecordCount = 0 Then
                MsgBox "Descupe-nos, não temos nenhuma função cadastrada", vbInformation, "Arbimy manager 2.0"
                Exit Sub
            End If
                mfg_funcoes.Rows = 2
                mfg_funcoes = Clear
                mfg_funcoes.FormatString = "Código  |Função        "
            tabfuncao.MoveFirst
                
            Do While tabfuncao.EOF = False
        
                mfg_funcoes.TextMatrix(mfg_funcoes.Rows - 1, 0) = tabfuncao!Codigo
                mfg_funcoes.TextMatrix(mfg_funcoes.Rows - 1, 1) = tabfuncao!Funcao
                tabfuncao.MoveNext
                mfg_funcoes.Rows = mfg_funcoes.Rows + 1
            Loop
                mfg_funcoes.Rows = mfg_funcoes.Rows - 1
End Sub
Private Sub alterar()
            tabfuncao!Codigo = txt_codigo.Text
            tabfuncao!Funcao = txt_funcao.Text
            tabfuncao.Update
End Sub
Private Sub Mostrar()
            txt_codigo.Text = tabfuncao!Codigo
            txt_funcao.Text = tabfuncao!Funcao
End Sub
