VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{D9D9230C-4947-4B31-8632-942172F30553}#1.1#0"; "BARCODE_ACTIVEX.ocx"
Begin VB.Form frm_codigodebarra 
   Caption         =   "CADASTRO DE CODIGO DE BARRA"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "CÓDIGO DE BARRA"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7335
      Begin VB.TextBox txt_codigo_barra 
         Height          =   285
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   12
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txt_codigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "CÓDIGO DE BARRA"
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "CÓDIGO"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
   End
   Begin BARCODE_ACTIVEX.cBarCode cbc_codigo 
      Height          =   1980
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   3493
      CodeName        =   "3"
      Value           =   ""
   End
   Begin VB.Frame Frame1 
      Caption         =   "CONTROLE"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   7335
      Begin VB.CommandButton cmd_imprimir 
         Caption         =   "IMPRIMIR"
         Height          =   375
         Left            =   5880
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_excluir 
         Caption         =   "EXCLUIR"
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmd_novo 
         Caption         =   "NOVO"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmd_alterar 
         Caption         =   "ALTERAR"
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmd_salvar 
         Caption         =   "SALVAR"
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid mfg_codigo_barrra 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3625
      _Version        =   393216
      FormatString    =   "CÓDIGO             | CÓDIGO DE BARRA                                 "
   End
End
Attribute VB_Name = "frm_codigodebarra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l_cod_codigo As Integer
Dim L_linha As Integer
Dim var As Integer

Private Sub cmd_alterar_Click()
            status = "alteradas"
            If txt_codigo_barra <> "" Then
            var = 0
            If tabcodigobarra.State = adStateOpen Then tabcodigobarra.Close
            tabcodigobarra.Open "select * from Codigo_barra where codigo=" & txt_codigo
            If tabcodigobarra.RecordCount <> 0 Then
            Call gravar_codigobarra
            If var = 0 Then Call box
            Call flex
            End If
            Else
            MsgBox "DIGITE O CÓDIGO DE BARRA", vbInformation, "TRICON SUPERMERCADOS LTDA."
            txt_codigo_barra.SetFocus
            End If
End Sub

Private Sub cmd_excluir_Click()
            status = "excluídas"
            If MsgBox("Deseja realmente excluir este código de barra ?", vbYesNo, "TRICON SUPERMERCADOS LTDA.") = vbYes Then
            If tabcodigobarra.State = adStateOpen Then tabcodigobarra.Close
            tabcodigobarra.Open "Select * from Codigo_barra where Codigo=" & txt_codigo
            If tabcodigobarra.RecordCount <> 0 Then
            conectar.Execute "Delete from Codigo_barra where Codigo=" & txt_codigo
            Call box
            Call flex
            cmd_novo_Click
            Else
            MsgBox "ESTE CÓDIGO DE BARRA NÃO EXISTE", vbInformation, "TRICON SUPERMERCADOS LTDA."
            End If
            End If
End Sub

Private Sub cmd_imprimir_Click()
cbc_codigo.PrintPicture
End Sub

Private Sub cmd_novo_Click()
cbc_codigo.Value = ""
Call cod_codigo
txt_codigo_barra = Clear
txt_codigo_barra.SetFocus
End Sub

Private Sub cmd_salvar_Click()
            status = "gravadas"
            If txt_codigo_barra <> "" Then
            var = 0
            If tabcodigobarra.State = adStateOpen Then tabcodigobarra.Close
            tabcodigobarra.Open "select * from Codigo_barra where codigo=" & txt_codigo
            If tabcodigobarra.RecordCount = 0 Then
            Call gravar_codigobarra
            If var = 0 Then Call box
            Call flex
            End If
            Else
            MsgBox "DIGITE O CÓDIGO DE BARRA", vbInformation, "TRICON SUPERMERCADOS LTDA."
            txt_codigo_barra.SetFocus
            End If
End Sub

Private Sub Form_Load()
cbc_codigo.Value = ""
Call abrir_banco
If tabcodigobarra.State = adStateOpen Then tabcodigobarra.Close
tabcodigobarra.Open "Codigo_barra", conectar, adOpenKeyset, adLockOptimistic
Call cod_codigo
Call flex
End Sub

Private Sub txt_codigo_barra_Change()
cbc_codigo.Value = txt_codigo_barra.Text
End Sub

Private Sub cod_codigo()
            l_cod_codigo = 1
A:
            If tabcodigobarra.State = adStateOpen Then tabcodigobarra.Close
            tabcodigobarra.Open "select * from Codigo_barra where Codigo =" & l_cod_codigo
            If tabcodigobarra.RecordCount > 0 Then
                l_cod_codigo = l_cod_codigo + 1
                GoTo A:
            End If
            txt_codigo = l_cod_codigo
End Sub

Private Sub gravar_codigobarra()
            If status <> "alteradas" Then
            tabcodigobarra.AddNew
            Else
            If tabcodigobarra.State = adStateOpen Then tabcodigobarra.Close
            tabcodigobarra.Open "select * from Codigo_barra where Codigo=" & txt_codigo
            If tabcodigobarra.RecordCount = 0 Then
            MsgBox "ESTE CÓDIGO DE BARRA NÃO EXISTE", vbInformation, "TRICON SUPERMERCADOS LTDA."
            var = 1
            Exit Sub
            End If
            End If
            tabcodigobarra!Codigo = txt_codigo
            tabcodigobarra!codigo_barra = txt_codigo_barra
            tabcodigobarra.Update
End Sub




Private Sub flex()
If tabcodigobarra.State = adStateOpen Then tabcodigobarra.Close
tabcodigobarra.Open "Codigo_barra", conectar, adOpenKeyset, adLockOptimistic
            
            If tabcodigobarra!Codigo <> "" Then tabcodigobarra.MoveFirst
            mfg_codigo_barrra.Clear
            mfg_codigo_barrra.Rows = 2
            mfg_codigo_barrra.FormatString = "CÓDIGO             | CÓDIGO DE BARRA                                 "
            Do Until tabcodigobarra.EOF = True
            mfg_codigo_barrra.TextMatrix(mfg_codigo_barrra.Rows - 1, 0) = tabcodigobarra!Codigo
            mfg_codigo_barrra.TextMatrix(mfg_codigo_barrra.Rows - 1, 1) = tabcodigobarra!codigo_barra
            mfg_codigo_barrra.Rows = mfg_codigo_barrra.Rows + 1
            tabcodigobarra.MoveNext
            Loop
            mfg_codigo_barrra.Rows = mfg_codigo_barrra.Rows - 1
End Sub

Private Sub mfg_codigo_barrra_Click()

If tabcodigobarra.State = adStateOpen Then tabcodigobarra.Close
tabcodigobarra.Open "Codigo_barra", conectar, adOpenKeyset, adLockOptimistic

            If mfg_codigo_barrra.Rows < 2 Then Exit Sub
            L_linha = mfg_codigo_barrra.Row
            l_cod_codigo = mfg_codigo_barrra.TextMatrix(L_linha, 0)
            If tabcodigobarra.State = adStateOpen Then tabcodigobarra.Close
            tabcodigobarra.Open "Select * From Codigo_barra Where Codigo = " & l_cod_codigo
            txt_codigo_barra = tabcodigobarra!codigo_barra
            txt_codigo = tabcodigobarra!Codigo
End Sub

Private Sub txt_codigo_barra_LostFocus()
            txt_codigo_barra.Text = UCase(txt_codigo_barra)
End Sub
