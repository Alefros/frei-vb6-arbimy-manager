VERSION 5.00
Begin VB.Form frm_login 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Login"
   ClientHeight    =   1110
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   3045
   Icon            =   "frm_login.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1110
   ScaleWidth      =   3045
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_senha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmd_entrar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Entrar"
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
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txt_login 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Senha"
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
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Login"
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
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim login, senha As String
Public tabusuarios As New ADODB.Recordset
Private Sub cmd_entrar_Click()
''''''''''''verifica se Login/Senha estão vazios''''''''''''''''''''''''''''''''
            If txt_login = Empty Then MsgBox "Digite seu login", vbInformation, "Arbimy manager 2.0"
            If txt_login = Empty Then Exit Sub
            If txt_senha = Empty Then MsgBox "Digite sua senha", vbInformation, "Arbimy manager 2.0"
            If txt_senha = Empty Then Exit Sub
'''''''''''verifica Login e senha e permite ou não o acesso ao sistema'''''''''
            Call fechar
            tabusuarios.Open "select * from Usuarios where Login = '" & txt_login.Text & "'"
            If tabusuarios.RecordCount = 0 Then GoTo A:
            If tabusuarios.RecordCount = 1 Then
            login = tabusuarios!login
            Call fechar
            tabusuarios.Open "select * from Usuarios where Login = '" & login & "'"
            senha = tabusuarios!senha
            'If tabusuarios.RecordCount = 0 Then
             '   txt_login = Clear
              '  txt_senha = Clear
               ' txt_login.SetFocus
            'MsgBox "O login e ou senha digitados são incorretos, favor verificar!", vbInformation, "Arbimy manager 2.0"
            'Exit Sub
            'End If
            
            If txt_senha = senha Then 'MDIForm1.Show
            frmSplash.Show
            usuario = txt_login
            usuario = UCase(usuario)
            
            frmSplash.lbl_bemvindo.Caption = " Bem vindo(a) Sr(a) " & usuario & ""
            Unload Me
            ElseIf txt_senha <> senha Then GoTo A:
                    
            End If
            End If
            Exit Sub
A:                                  MsgBox "O login e ou senha digitados são incorretos, favor verificar!", vbInformation, "Arbimy manager 2.0"
            txt_senha = Clear
            txt_login = Clear
            txt_login.SetFocus
End Sub
Private Sub abrir()
            Call fechar
            tabusuarios.Open "Usuarios", conectar, adOpenKeyset, adLockOptimistic
End Sub
Private Sub fechar()
            If tabusuarios.State = 1 Then tabusuarios.Close
End Sub

Private Sub Form_Load()
            Call abrir_banco
            Call abrir
End Sub
Private Sub entrar()

End Sub
Private Sub txt_senha_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
        Call cmd_entrar_Click
        End If
End Sub
