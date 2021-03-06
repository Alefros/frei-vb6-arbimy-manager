VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_contas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contas Correntes"
   ClientHeight    =   4785
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   6090
   Icon            =   "frm_contas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSFlexGridLib.MSFlexGrid mfg_contas 
      Height          =   1935
      Left            =   0
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2760
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3413
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Caption         =   "COMANDOS"
      Height          =   975
      Left            =   0
      TabIndex        =   16
      Top             =   1680
      Width           =   6015
      Begin VB.CommandButton cmd_novo 
         Height          =   615
         Left            =   720
         Picture         =   "frm_contas.frx":1CCA
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Novo"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmd_salvar 
         Height          =   615
         Left            =   1920
         Picture         =   "frm_contas.frx":1FD4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salvar"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmd_alterar 
         Height          =   615
         Left            =   3240
         Picture         =   "frm_contas.frx":22DE
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Alterar"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmd_excluir 
         Height          =   615
         Left            =   4560
         Picture         =   "frm_contas.frx":25E8
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   240
         Width           =   615
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2990
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "BANCOS"
      TabPicture(0)   =   "frm_contas.frx":28F2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_banco"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txt_banco"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txt_codigo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "AG�NCIAS"
      TabPicture(1)   =   "frm_contas.frx":290E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "lbl_agencia"
      Tab(1).Control(2)=   "msk_agencia"
      Tab(1).Control(3)=   "cbo_banco"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "CONTAS CORRENTES"
      TabPicture(2)   =   "frm_contas.frx":292A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label6"
      Tab(2).Control(1)=   "Label3"
      Tab(2).Control(2)=   "cbo_agencia"
      Tab(2).Control(3)=   "msk_conta"
      Tab(2).ControlCount=   4
      Begin MSMask.MaskEdBox msk_conta 
         Height          =   255
         Left            =   -73680
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   11
         Mask            =   "99.999999-9"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_codigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txt_banco 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   1080
         Width           =   3615
      End
      Begin VB.ComboBox cbo_banco 
         Height          =   315
         Left            =   -73680
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox cbo_agencia 
         Height          =   315
         Left            =   -73680
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin MSMask.MaskEdBox msk_agencia 
         Height          =   255
         Left            =   -73680
         TabIndex        =   4
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "9999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "C�DIGO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lbl_banco 
         Caption         =   "BANCO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lbl_agencia 
         Caption         =   "AGENCIA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   14
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "CONTA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   13
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "BANCO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   12
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "AG�NCIA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "frm_contas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L_linha As Integer
Dim var As String
Dim codbanco As String
Dim codagencia As String
Dim l_codbanco As String
Dim l_codagencia As Integer

Private Sub cbo_agencia_Change()
            If cbo_agencia.ListIndex <> -1 Then l_codbanco = cbo_agencia.ItemData(cbo_agencia.ListIndex)
End Sub

Private Sub cbo_banco_Change()
            If cbo_banco.ListIndex <> -1 Then l_codbanco = cbo_banco.ItemData(cbo_banco.ListIndex)
End Sub

Private Sub cmd_alterar_Click()
            status = "Alteradas"
Call Desabilitar
            If SSTab1.Tab = 0 Then
            If txt_banco = "" Then MsgBox "DIGITE UM BANCO V�LIDO", vbInformation, "TRICON SUPERMERCADOS LTDA.": txt_banco.SetFocus: Exit Sub
            ElseIf SSTab1.Tab = 1 Then
            If cbo_banco = "" Then MsgBox "SELECIONE UM BANCO V�LIDO", vbInformation, "TRICON SUPERMERCADOS LTDA.": cbo_banco.SetFocus: Exit Sub
            If msk_agencia = "" Then MsgBox "DIGITE UMA AGENCIA V�LIDA", vbInformation, "TRICON SUPERMERCADOS LTDA.": msk_agencia.SetFocus: Exit Sub
            ElseIf SSTab1.Tab = 2 Then
            If cbo_agencia = "" Then MsgBox "SELECIONE UMA AGENCIA V�LIDA", vbInformation, "TRICON SUPERMERCADOS LTDA.": cbo_agencia.SetFocus: Exit Sub
            If msk_conta = "" Then MsgBox "DIGITE UMA CONTA CORRENTE V�LIDA", vbInformation, "TRICON SUPERMERCADOS LTDA.": msk_conta.SetFocus: Exit Sub
            End If
            Call gravar_contas
            If var = 0 Then Call box
Call Habilitar
End Sub

Private Sub cmd_excluir_Click()
status = "Exclu�das"


    If SSTab1.Tab = 0 Then
            On Error GoTo A:
        If MsgBox("DESEJA REALMENTE EXCLUIR ESTE BANCO ?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
                If tabbanco.State = adStateOpen Then tabbanco.Close
                tabbanco.Open "Select * from Bancos where Banco ='" & txt_banco & "'"
                    If tabbanco.RecordCount <> 0 Then
                    conectar.Execute "delete from Bancos where Banco = '" & txt_banco & "'"
                    Call flex
                    Call box
                    Call cod_contas
                    txt_banco = Clear
                    txt_banco.SetFocus
                    Else
                    MsgBox "ESTE BANCO N�O EXISTE", vbInformation, "TRICON SUPERMERCADOS LTDA."
                    txt_banco.SetFocus
                    End If
        End If
        Exit Sub
A:
            If Err.Description = "O registro n�o pode ser exclu�do ou alterado porque a tabela 'Agencias' inclui registros relacionados a ele." Then
            MsgBox "EXISTE UM REGISTRO NAS AGENCIAS RELACIONADO COM ESTE BANCO" & vbCrLf & "EXCLUA O REGISTRO E TENTE NOVAMENTE", vbInformation, "TRICON SUPERMERCADOS LTDA."
            SSTab1.Tab = 1
            End If
            
    ElseIf SSTab1.Tab = 1 Then
            On Error GoTo B:
            Call Desabilitar
        If MsgBox("DESEJA REALMENTE EXCLUIR ESTA AGENCIA ?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
                If tabagencias.State = adStateOpen Then tabagencias.Close
                tabagencias.Open "Select * from Agencias where Agencia = '" & msk_agencia & "'"
                    If tabagencias.RecordCount <> 0 Then
                    conectar.Execute "delete from Agencias where Agencia = '" & msk_agencia & "'"
                    Call flex
                    Call box
                    Call cod_contas
                    cbo_banco.Text = ""
                    msk_agencia.Text = ""
                    cbo_banco.SetFocus
                    Else
                    MsgBox "ESTA AGENCIA N�O EXISTE", vbInformation, "TRICON SUPERMERCADOS LTDA."
                    cbo_banco.SetFocus
                    End If
        End If
            Call Habilitar
            Exit Sub
B:
            If Err.Description = "O registro n�o pode ser exclu�do ou alterado porque a tabela 'Contas_correntes' inclui registros relacionados a ele." Then
            MsgBox "EXISTE UM REGISTRO NAS CONTAS CORRENTES RELACIONADO COM ESTA AGENCIA" & vbCrLf & "EXCLUA O REGISTRO E TENTE NOVAMENTE", vbInformation, "TRICON SUPERMERCADOS LTDA."
            SSTab1.Tab = 2
            End If
    ElseIf SSTab1.Tab = 2 Then
    Call Desabilitar
        If MsgBox("DESEJA REALMENTE EXCLUIR ESTA CONTA CORRENTE ?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
                If tabconta.State = adStateOpen Then tabconta.Close
                tabconta.Open "Select * from Contas_correntes where Conta_corrente ='" & msk_conta & "'"
                    If tabconta.RecordCount <> 0 Then
                    conectar.Execute "delete from Contas_correntes where Conta_corrente = '" & msk_conta & "'"
                    Call flex
                    Call box
                    Call cod_contas
                    cbo_agencia.Text = ""
                    msk_conta.Text = ""
                    cbo_agencia.SetFocus
                    Else
                    MsgBox "ESTA CONTA CORRENTE N�O EXISTE", vbInformation, "TRICON SUPERMERCADOS LTDA."
                    cbo_agencia.SetFocus
                    End If
    End If
Call Habilitar
    End If
End Sub
Function cod_contas()
            l_codbanco = 1
A:
            If tabbanco.State = adStateOpen Then tabbanco.Close
            tabbanco.Open "select * from Bancos where cod_banco = " & l_codbanco
            If tabbanco.RecordCount > 0 Then
                l_codbanco = l_codbanco + 1
                GoTo A
            End If
            txt_codigo = l_codbanco
End Function

Private Sub limpar()
txt_banco = Clear
Call Habilitar
msk_agencia = Clear
Call Desabilitar
txt_conta = Clear
End Sub

Private Sub cmd_novo_Click()
        Call Desabilitar
            If SSTab1.Tab = 0 Then
                txt_codigo = Clear
                txt_banco = Clear
                txt_banco.SetFocus
                Call cod_contas
            ElseIf SSTab1.Tab = 1 Then
                cbo_banco.Text = ""
                msk_agencia.Text = ""
                cbo_banco.SetFocus
            ElseIf SSTab1.Tab = 2 Then
                cbo_agencia.Text = ""
                msk_conta = ""
                cbo_agencia.SetFocus
            End If
        Call Habilitar
End Sub

Private Sub cmd_salvar_Click()
            status = "Gravadas"
                '////////////////////////////SSTAB 1\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Call Desabilitar
            If SSTab1.Tab = 0 Then
            If txt_banco = "" Then MsgBox "DIGITE UM BANCO V�LIDO", vbInformation, "TRICON SUPERMERCADOS LTDA.": txt_banco.SetFocus: Exit Sub
            ElseIf SSTab1.Tab = 1 Then
            If cbo_banco = "" Then MsgBox "SELECIONE UM BANCO V�LIDO", vbInformation, "TRICON SUPERMERCADOS LTDA.": cbo_banco.SetFocus: Exit Sub
            If msk_agencia = "" Then MsgBox "DIGITE UMA AGENCIA V�LIDA", vbInformation, "TRICON SUPERMERCADOS LTDA.": msk_agencia.SetFocus: Exit Sub
            ElseIf SSTab1.Tab = 2 Then
            If cbo_agencia = "" Then MsgBox "SELECIONE UMA AGENCIA V�LIDA", vbInformation, "TRICON SUPERMERCADOS LTDA.": cbo_agencia.SetFocus: Exit Sub
            If msk_conta = "" Then MsgBox "DIGITE UMA CONTA CORRENTE V�LIDA", vbInformation, "TRICON SUPERMERCADOS LTDA.": msk_conta.SetFocus: Exit Sub
            End If
            
    If SSTab1.Tab = 0 Then
                Call gravar_contas
                If var = 0 Then Call box
                Call fechar
                Call abrir
                Call flex
    
               '////////////////////////////SSTAB 2\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
            ElseIf SSTab1.Tab = 1 Then
                Call gravar_contas
                If var = 0 Then Call box
                If var = 0 Then Call fechar
                If var = 0 Then Call abrir
                Call flex
             '////////////////////////////SSTAB 3\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            ElseIf SSTab1.Tab = 2 Then
                Call gravar_contas
                If var = 0 Then Call box
                If var = 0 Then Call fechar
                If var = 0 Then Call abrir
                Call flex
    End If
Call Habilitar
End Sub

Private Sub flex()
            If tabbanco.State = adStateOpen Then tabbanco.Close
            tabbanco.Open "Bancos", conectar, adOpenKeyset, adLockOptimistic
            
            If tabagencias.State = adStateOpen Then tabagencias.Close
            tabagencias.Open "Agencias", conectar, adOpenKeyset, adLockOptimistic
            
            If tabconta.State = adStateOpen Then tabconta.Close
            tabconta.Open "Contas_correntes", conectar, adOpenKeyset, adLockOptimistic
            
            
            mfg_contas.Clear
            mfg_contas.Rows = 2
                        
            
            If SSTab1.Tab = 0 Then Call flex_bancos
            If SSTab1.Tab = 1 Then Call flex_agencias
            If SSTab1.Tab = 2 Then Call flex_contas_correntes
            
            mfg_contas.Rows = mfg_contas.Rows - 1
            
End Sub
Private Sub gravar_agencia()
            If tabbanco.State = adStateOpen Then tabbanco.Close
            tabbanco.Open "select * from Bancos where Banco = '" & cbo_banco & "'"
            codbanco = tabbanco!Cod_banco
            If status <> "Alteradas" Then
            tabagencias.AddNew
            Else
            If tabagencias.State = adStateOpen Then tabagencias.Close
            tabagencias.Open "Select * From Agencias where Agencia =" & msk_agencia
            End If
            tabagencias!Agencia = msk_agencia
            tabagencias!Cod_banco = codbanco
            tabagencias.Update
            Call flex
            Call Carregar_Combo
            Call cod_contas
End Sub
Function Habilitar()
            msk_agencia.PromptInclude = True
            msk_conta.PromptInclude = True
End Function
Function Desabilitar()
            msk_agencia.PromptInclude = False
            msk_conta.PromptInclude = False
End Function
Private Sub gravar_contas()
            Call Desabilitar
            
            If tabbanco.State = adStateOpen Then tabbanco.Close
            tabbanco.Open "Bancos", conectar, adOpenKeyset, adLockOptimistic
            
            If tabagencias.State = adStateOpen Then tabagencias.Close
            tabagencias.Open "Agencias", conectar, adOpenKeyset, adLockOptimistic
            
            If tabconta.State = adStateOpen Then tabconta.Close
            tabconta.Open "Contas_correntes", conectar, adOpenKeyset, adLockOptimistic
            
            
            If SSTab1.Tab = 0 Then
            'On Error GoTo A:
            
            If status <> "Alteradas" Then
            If tabbanco.State = adStateOpen Then tabbanco.Close
            tabbanco.Open "Select * from Bancos Where Cod_banco =" & txt_codigo
                If tabbanco.RecordCount <> 0 Then
                MsgBox "ESTE BANCO J� EXISTE", vbInformation, "TRICON SUPERMERCADOS LTDA."
                var = 1
                Exit Sub
                End If
            tabbanco.AddNew
            tabbanco!Cod_banco = txt_codigo
            tabbanco!Banco = txt_banco
            tabbanco.Update
            Else
            If tabbanco.State = adStateOpen Then tabbanco.Close
            tabbanco.Open "Select * from Bancos Where Cod_banco =" & txt_codigo
                If tabbanco.RecordCount = 0 Then
                MsgBox "ESTE BANCO N�O EXISTE", vbInformation, "TRICON SUPERMERCADOS LTDA."
                var = 1
                Exit Sub
                End If
            tabbanco!Banco = txt_banco.Text
            tabbanco.Update
            GoTo CARREGAR:
            End If
        
            tabbanco!Cod_banco = txt_codigo
            tabbanco!Banco = txt_banco.Text
            tabbanco.Update
            GoTo CARREGAR:
            
A:
            If Err.Description = "O registro n�o pode ser exclu�do ou alterado porque a tabela 'Agencias' inclui registros relacionados a ele." Then
            MsgBox "EXISTE UM REGISTRO NAS AGENCIAS RELACIONADO COM ESTE BANCO" & vbCrLf & "EXCLUA O REGISTRO E TENTE NOVAMENTE", vbInformation, "TRICON SUPERMERCADOS LTDA."
            SSTab1.Tab = 1
            End If
            
            
            ElseIf SSTab1.Tab = 1 Then
            On Error GoTo B:
            If status <> "Alteradas" Then
            If tabagencias.State = adStateOpen Then tabagencias.Close
            tabagencias.Open "Select * from Agencias Where Agencia = '" & msk_agencia & "'"
                If tabagencias.RecordCount <> 0 Then
                MsgBox "ESTA AGENCIA J� EXISTE", vbInformation, "TRICON SUPERMERCADOS LTDA."
                var = 1
                Exit Sub
                End If
            tabagencias.AddNew
            tabagencias!Agencia = msk_agencia
            tabagencias!Cod_banco = cbo_banco.ListIndex + 1
            tabagencias.Update
            GoTo CARREGAR:
            Else
            If tabagencias.State = adStateOpen Then tabagencias.Close
            tabagencias.Open "Select * from Agencias Where Agencia = '" & msk_agencia & "'"
                If tabagencias.RecordCount = 0 Then
                MsgBox "ESTA AGENCIA N�O EXISTE", vbInformation, "TRICON SUPERMERCADOS LTDA."
                var = 1
                Exit Sub
                End If
            tabagencias!Cod_banco = cbo_banco.ListIndex + 1
            tabagencias.Update
            GoTo CARREGAR:
            End If
            
B:
            If Err.Description = "O registro n�o pode ser exclu�do ou alterado porque a tabela 'Contas_correntes' inclui registros relacionados a ele." Then
            MsgBox "EXISTE UM REGISTRO NAS CONTAS CORRENTES RELACIONADO COM ESTA AGENCIA" & vbCrLf & "EXCLUA O REGISTRO E TENTE NOVAMENTE", vbInformation, "TRICON SUPERMERCADOS LTDA."
            SSTab1.Tab = 2
            End If
            ElseIf SSTab1.Tab = 2 Then
            If status <> "Alteradas" Then
            If tabconta.State = adStateOpen Then tabconta.Close
            tabconta.Open "Select * from Contas_correntes Where Conta_corrente = '" & msk_conta & "'"
                If tabconta.RecordCount <> 0 Then
                MsgBox "ESTA CONTA CORRENTE J� EXISTE", vbInformation, "TRICON SUPERMERCADOS LTDA."
                var = 1
                Exit Sub
                End If
            tabconta.AddNew
            tabconta!Conta_corrente = msk_conta
            tabconta!Agencia = cbo_agencia.ItemData(cbo_agencia.ListIndex)
            tabconta.Update
            GoTo CARREGAR:
            Else
            If tabconta.State = adStateOpen Then tabconta.Close
            tabconta.Open "Select * from Contas_correntes Where Conta_corrente = '" & msk_conta & "'"
                If tabconta.RecordCount = 0 Then
                MsgBox "ESTA CONTA CORRENTE N�O EXISTE", vbInformation, "TRICON SUPERMERCADOS LTDA."
                var = 1
                Exit Sub
                End If
            tabconta!Agencia = cbo_agencia.ItemData(cbo_agencia.ListIndex)
            tabconta.Update
            End If
            End If
CARREGAR:
            var = 0
            Call flex
            Call Habilitar
End Sub

Private Sub Form_Load()
            Call abrir
            Call flex
            Call cod_contas
            Call Carregar_Combo
End Sub

Private Sub mfg_contas_Click()
            Call abrir
            Call Desabilitar
            If mfg_contas.Rows < 2 Then Exit Sub
            L_linha = mfg_contas.Row
            l_codbanco = mfg_contas.TextMatrix(L_linha, 0)
            
            
            If SSTab1.Tab = 0 Then
            
            If tabbanco.State = adStateOpen Then tabbanco.Close
            tabbanco.Open "Select * From Bancos Where Cod_banco = " & l_codbanco
            txt_banco = tabbanco!Banco
            txt_codigo = tabbanco!Cod_banco
            
            
            ElseIf SSTab1.Tab = 1 Then
            
            If tabagencias.State = adStateOpen Then tabagencias.Close
            tabagencias.Open "Select * From Agencias Where Agencia = '" & mfg_contas.TextMatrix(L_linha, 0) & "'"
            msk_agencia = tabagencias!Agencia
            
            If tabbanco.State = adStateOpen Then tabbanco.Close
            tabbanco.Open "Select * From Bancos Where Cod_banco = " & tabagencias!Cod_banco
            cbo_banco.Text = tabbanco!Banco
            
            
            ElseIf SSTab1.Tab = 2 Then
            
            msk_conta = mfg_contas.TextMatrix(L_linha, 0)
            cbo_agencia = mfg_contas.TextMatrix(L_linha, 1)
            
            End If
            Call Habilitar
End Sub
Private Sub Carregar_Combo()
            
            
            cbo_banco.Clear
            cbo_agencia.Clear
            If tabbanco.EOF = True Then Call abrir_banco: tabbanco.Open "Bancos", conectar, adOpenKeyset, adLockOptimistic
            If tabagencias.State = adStateClose Then tabagencias.Open "Agencias", conectar, adOpenKeyset, adLockOptimistic
            If tabagencias.EOF = True Then Call abrir_banco: tabagencias.Open "Agencias", conectar, adOpenKeyset, adLockOptimistic
            
            If tabbanco.State = adStateClose Then tabbanco.Open "Bancos", conectar, adOpenKeyset, adLockOptimistic
            If tabagencias.State = adStateClose Then tabagencias.Open "Agencias", conectar, adOpenKeyset, adLockOptimistic
            If tabbanco.EOF = False Then tabbanco.MoveFirst
            If tabagencias.EOF = False Then tabagencias.MoveFirst
            
            Do Until tabbanco.EOF = True
            cbo_banco.AddItem tabbanco!Banco
            cbo_banco.ItemData(cbo_banco.NewIndex) = tabbanco!Cod_banco
            tabbanco.MoveNext
            Loop
            
            
            Do Until tabagencias.EOF = True
            cbo_agencia.AddItem tabagencias!Agencia
            cbo_agencia.ItemData(cbo_agencia.NewIndex) = tabagencias!Agencia
            tabagencias.MoveNext
            Loop
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
            Call flex
            Call Carregar_Combo
End Sub
Private Sub flex_bancos()
            mfg_contas.FormatString = "C�digo |Banco                           "
            Do Until tabbanco.EOF = True
            mfg_contas.TextMatrix(mfg_contas.Rows - 1, 0) = tabbanco!Cod_banco
            mfg_contas.TextMatrix(mfg_contas.Rows - 1, 1) = tabbanco!Banco
            mfg_contas.Rows = mfg_contas.Rows + 1
            tabbanco.MoveNext
            Loop
End Sub
Private Sub flex_agencias()
            mfg_contas.FormatString = "Agencia |Banco                           "
            Do Until tabagencias.EOF = True
            If tabbanco.State = adStateOpen Then tabbanco.Close
            tabbanco.Open "select * from Bancos where Cod_banco =" & tabagencias!Cod_banco
            codbanco = tabbanco!Banco
            mfg_contas.TextMatrix(mfg_contas.Rows - 1, 0) = tabagencias!Agencia
            mfg_contas.TextMatrix(mfg_contas.Rows - 1, 1) = codbanco
            mfg_contas.Rows = mfg_contas.Rows + 1
            tabagencias.MoveNext
            Loop
End Sub
Private Sub flex_contas_correntes()
            mfg_contas.FormatString = "Conta Corrente  |Agencia                      "
            Do Until tabconta.EOF = True
            mfg_contas.TextMatrix(mfg_contas.Rows - 1, 0) = tabconta!Conta_corrente
            mfg_contas.TextMatrix(mfg_contas.Rows - 1, 1) = tabconta!Agencia
            mfg_contas.Rows = mfg_contas.Rows + 1
            tabconta.MoveNext
            Loop
End Sub
