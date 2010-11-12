VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_formaspagamento 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formas de pagamento"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4335
   Icon            =   "frm_formaspagamento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4335
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Comandos"
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   2280
      Width           =   4335
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
         Left            =   360
         Picture         =   "frm_formaspagamento.frx":1CCA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Novo"
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
         Left            =   1320
         Picture         =   "frm_formaspagamento.frx":1FD4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salvar"
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
         Left            =   2280
         Picture         =   "frm_formaspagamento.frx":22DE
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Alterar"
         Top             =   240
         Width           =   735
      End
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
         Left            =   3240
         Picture         =   "frm_formaspagamento.frx":25E8
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   240
         Width           =   735
      End
   End
   Begin MSFlexGridLib.MSFlexGrid mfg_pagamento 
      Height          =   1575
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2778
      _Version        =   393216
      BackColorFixed  =   14737632
      FormatString    =   "Código           |           Forma de pagamento  "
   End
   Begin VB.TextBox txt_pagamento 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox txt_codigo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pagamento"
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
      Width           =   1215
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
      Width           =   855
   End
End
Attribute VB_Name = "frm_formaspagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L_Colunas, L_Linha As Long
Dim L_Codpagamento As Integer
Public Codpagamento As New ADODB.Recordset

Private Sub cmd_alterar_Click()
            status = "alteradas"
            If tabpagamento.State = adStateOpen Then tabpagamento.Close
                tabpagamento.Open "Select * From Formas_pagamentos where Codigo like '" & txt_codigo & "'"
                If tabpagamento.RecordCount <> 0 Then
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
                If tabpagamento.State = adStateOpen Then tabpagamento.Close
                    tabpagamento.Open "select * from Formas_pagamentos where Codigo = " & txt_codigo
                    If tabpagamento.RecordCount <> 0 Then
                        conectar.Execute "Delete From Formas_pagamentos where Codigo like '" & txt_codigo & "'"
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
            If txt_pagamento = Empty Then
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
            Call cod_pagamento
            Call carregar_lista
End Sub

Private Sub Form_Unload(Cancel As Integer)
            Call voltar_botao
End Sub

Private Sub mfg_pagamento_Click()
            L_Linha = mfg_pagamento.Row
            L_Codpagamento = mfg_pagamento.TextMatrix(L_Linha, 0)
            Call abrir
            tabpagamento.Close
            tabpagamento.Open "select * from Formas_pagamentos where Codigo = " & L_Codpagamento
            Call Mostrar
End Sub

Private Sub txt_pagamento_LostFocus()
            cmd_salvar.Enabled = True
End Sub
Private Sub limpar()
            txt_codigo = Clear
            txt_pagamento = Clear
            txt_pagamento.SetFocus
            cod_pagamento
End Sub
Private Sub gravar()
            Call abrir
            tabpagamento.Close
            tabpagamento.Open "select * from Formas_pagamentos where Codigo = " & txt_codigo
            If tabpagamento.RecordCount = 0 Then
            If status <> "alteradas" Then Codpagamento.AddNew
            Codpagamento!Codigo = txt_codigo.Text
            Codpagamento!Forma_pagamento = txt_pagamento.Text
            Codpagamento.Update
            End If
End Sub
Private Sub abrir()
            If tabpagamento.State = adStateOpen Then tabpagamento.Close
            tabpagamento.Open "Formas_pagamentos", conectar, adOpenKeyset, adLockOptimistic
            If Codpagamento.State = adStateOpen Then Codpagamento.Close
            Codpagamento.Open "Formas_pagamentos", conectar, adOpenKeyset, adLockOptimistic
End Sub
Private Sub cod_pagamento()
            Dim codpag As Integer
            codpag = 1
A:
            Call abrir
            If Codpagamento.State = adStateOpen Then Codpagamento.Close
            Codpagamento.Open "select * from Formas_pagamentos where Codigo = " & codpag
            If Codpagamento.RecordCount > 0 Then
                codpag = codpag + 1
                GoTo A
            End If
            txt_codigo = codpag
End Sub
Private Sub carregar_lista()
            Call abrir
            If tabpagamento.State = adStateOpen Then tabpagamento.Close
            tabpagamento.Open "select * from Formas_pagamentos"
            If tabpagamento.RecordCount = 0 Then
                MsgBox "Descupe-nos, não temos nenhuma forma de pagameto cadastrada", vbInformation, "Arbimy manager 2.0"
                Exit Sub
            End If
                mfg_pagamento.Rows = 2
                mfg_pagamento = Clear
                mfg_pagamento.FormatString = "Código  |Forma de pagamento  "
            tabpagamento.MoveFirst
            Do While tabpagamento.EOF = False
                mfg_pagamento.TextMatrix(mfg_pagamento.Rows - 1, 0) = tabpagamento!Codigo
                mfg_pagamento.TextMatrix(mfg_pagamento.Rows - 1, 1) = tabpagamento!Forma_pagamento
                tabpagamento.MoveNext
                mfg_pagamento.Rows = mfg_pagamento.Rows + 1
                Loop
                mfg_pagamento.Rows = mfg_pagamento.Rows - 1
End Sub
Private Sub Mostrar()
            txt_codigo.Text = tabpagamento!Codigo
            txt_pagamento.Text = tabpagamento!Forma_pagamento
            cmd_alterar.Enabled = True
            cmd_excluir.Enabled = True
End Sub
Private Sub alterar()
            tabpagamento!Forma_pagamento = txt_pagamento.Text
            tabpagamento.Update
            Call limpar
End Sub
