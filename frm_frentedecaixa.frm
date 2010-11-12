VERSION 5.00
Begin VB.Form frm_frentedecaixa 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option4 
      BackColor       =   &H00800000&
      Height          =   220
      Left            =   240
      MaskColor       =   &H00800000&
      TabIndex        =   20
      Top             =   10995
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00800000&
      Height          =   220
      Left            =   240
      MaskColor       =   &H00800000&
      TabIndex        =   19
      Top             =   10515
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00800000&
      Height          =   220
      Left            =   240
      MaskColor       =   &H00800000&
      TabIndex        =   18
      Top             =   10035
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      Height          =   220
      Left            =   240
      MaskColor       =   &H00800000&
      TabIndex        =   17
      Top             =   9540
      Width           =   255
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      ItemData        =   "frm_frentedecaixa.frx":0000
      Left            =   12720
      List            =   "frm_frentedecaixa.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   7800
      Width           =   2055
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      ItemData        =   "frm_frentedecaixa.frx":0004
      Left            =   10680
      List            =   "frm_frentedecaixa.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   7800
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      ItemData        =   "frm_frentedecaixa.frx":0008
      Left            =   6720
      List            =   "frm_frentedecaixa.frx":000A
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   7800
      Width           =   3975
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "OPERADOR DO CAIXA: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   6720
      TabIndex        =   22
      Top             =   11040
      Width           =   3180
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "     PAGAMENTO À VISTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   240
      TabIndex        =   21
      Top             =   9480
      Width           =   3360
   End
   Begin VB.Image Image2 
      Height          =   1470
      Left            =   13080
      Picture         =   "frm_frentedecaixa.frx":000C
      Stretch         =   -1  'True
      Top             =   9480
      Width           =   1800
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "     CANCELAR COMPRA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   240
      TabIndex        =   16
      Top             =   10920
      Width           =   3210
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "     CARTÃO DE DÉBITO   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   240
      TabIndex        =   15
      Top             =   10440
      Width           =   3360
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "     CARTÃO DE CRÉDITO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   240
      TabIndex        =   14
      Top             =   9960
      Width           =   3345
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PREÇO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   12720
      TabIndex        =   5
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTIDADE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10680
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6720
      TabIndex        =   1
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   6720
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   8055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   15270
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   15135
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PREÇO: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   6840
      TabIndex        =   13
      Top             =   9960
      Width           =   1140
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "QUANTIDADE: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   120
      TabIndex        =   12
      Top             =   8880
      Width           =   1995
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALIDADE: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   6840
      TabIndex        =   11
      Top             =   10560
      Width           =   1530
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MARCA: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   6840
      TabIndex        =   10
      Top             =   9360
      Width           =   1170
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "CÓDIGO DE BARRA: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   120
      TabIndex        =   8
      Top             =   8280
      Width           =   2760
   End
   Begin VB.Image Image1 
      Height          =   6135
      Left            =   120
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   6495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL: R$ 0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   11520
      TabIndex        =   7
      Top             =   8760
      Width           =   2580
   End
   Begin VB.Label lbl_subtotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUBTOTAL: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   6840
      TabIndex        =   6
      Top             =   8760
      Width           =   2100
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   6135
      Left            =   12720
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   6135
      Left            =   10680
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   6135
      Left            =   6720
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   6720
      Shape           =   4  'Rounded Rectangle
      Top             =   9360
      Width           =   6135
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   6720
      Shape           =   4  'Rounded Rectangle
      Top             =   9960
      Width           =   6135
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   6720
      Shape           =   4  'Rounded Rectangle
      Top             =   10560
      Width           =   6135
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   6720
      Shape           =   4  'Rounded Rectangle
      Top             =   8760
      Width           =   8295
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   6165
      Left            =   6960
      Top             =   2280
      Width           =   8055
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Left            =   120
      Top             =   9480
      Width           =   3735
   End
End
Attribute VB_Name = "frm_frentedecaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim codigo_barra, var, codigo As String
Private Sub Command1_Click()
            If List1.Top > 1560 And List1.ListCount <> 0 Then List1.Height = List1.Height + 360: List1.Top = List1.Top - 360
            List1.AddItem (Text1.Text), List1.ListCount
            List1.ListIndex = List1.ListCount - 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
            If KeyAscii = vbKeyEscape Then Unload Me
            If KeyAscii = 45 Then Call Remover
            If KeyAscii = 13 And Label6.Caption = "QUANTIDADE: " Then
            var = 1
            codigo_barra = ""
            Label6.BackColor = &H800000
            Label5.BackColor = &HFFFFFF
            Label5.ForeColor = &H0&
            Label6.ForeColor = &HFFFFFF
            End If
            
            If KeyAscii = 13 And Label6.Caption <> "QUANTIDADE: " And Image1.Picture <> 0 And var = 1 Then
            var = 0
            
            If List1.Top > 2041 And List1.ListCount <> 0 Then List1.Height = List1.Height + 360: List1.Top = List1.Top - 360
            List1.AddItem (Mid(Label7, 1, Len(Label7))), List1.ListCount
            List1.ListIndex = List1.ListCount - 1
            
            If List2.Top > 2041 And List2.ListCount <> 0 Then List2.Height = List2.Height + 360: List2.Top = List2.Top - 360
            List2.AddItem (Mid(Label6, 12, Len(Label6))), List2.ListCount
            List2.ListIndex = List2.ListCount - 1
            
            If List3.Top > 2041 And List3.ListCount <> 0 Then List3.Height = List3.Height + 360: List3.Top = List3.Top - 360
            List3.AddItem (Mid(lbl_subtotal, 11, Len(lbl_subtotal))), List3.ListCount
            List3.ListIndex = List3.ListCount - 1
            
            Label4.Caption = "TOTAL: " & Format(CCur(Mid(Label4, 7, Len(Label4))) + CCur(CCur(Mid(Label9, 7, Len(Label9))) * CCur(Mid(Label6, 12, Len(Label6)))), "currency")
            Label5.Caption = "CÓDIGO DE BARRA: "
            Label6.Caption = "QUANTIDADE: "
            lbl_subtotal.Caption = "SUBTOTAL: "
            Label7.Caption = ""
            Label8.Caption = "MARCA: "
            Label9.Caption = "PREÇO: "
            Label10.Caption = "VALIDADE: "
            Label6.BackColor = &HFFFFFF
            Label5.BackColor = &H800000
            Label5.ForeColor = &HFFFFFF
            Label6.ForeColor = &H0&
            
            ElseIf KeyAscii = 13 And Label7.Caption = "" Then
            var = 0
            Label5.Caption = "CÓDIGO DE BARRA: "
            Label6.Caption = "QUANTIDADE: "
            lbl_subtotal.Caption = "SUBTOTAL: "
            Label7.Caption = ""
            Label8.Caption = "MARCA: "
            Label9.Caption = "PREÇO: "
            Label10.Caption = "VALIDADE: "
            Label6.BackColor = &HFFFFFF
            Label5.BackColor = &H800000
            Label5.ForeColor = &HFFFFFF
            Label6.ForeColor = &H0&
            End If
            
'''''''''''''''''''''''''''DIGITAR NUMEROS'''''''''''''''''''''''''''''''''''''
            If KeyAscii > 47 And KeyAscii < 58 And var = 1 Then
            Label6 = Label6.Caption & Chr(KeyAscii)
            End If
            
            If KeyAscii > 47 And KeyAscii < 58 And var = 0 Then
            Label5 = Label5.Caption & Chr(KeyAscii)
            End If
'''''''''''''''''''''''''''DIGITAR NUMEROS'''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''DIGITAR LETRAS''''''''''''''''''''''''''''''''''''''
            If (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 123) Then
            If var = 0 Then Label5 = Label5.Caption & Chr(KeyAscii)
            End If
'''''''''''''''''''''''''''DIGITAR LETRAS''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''APAGAR CARACTERES'''''''''''''''''''''''''''''''''''
            If KeyAscii = 8 And var = 0 And Len(Label5) <> 17 Then
            Label5 = Left(Label5.Caption, Len(Label5) - 1)
            End If
            
            If KeyAscii = 8 And var = 1 And Len(Label6) <> 12 Then
            Label6 = Left(Label6.Caption, Len(Label6) - 1)
            End If
            
            If KeyAscii = 8 And var = 1 And Len(Label6) = 12 Then
            lbl_subtotal = "SUBTOTAL: "
            End If
''''''''''''''''''''''''''APAGAR CARACTERES''''''''''''''''''''''''''''''''''''
End Sub

Private Sub Remover()
If List1.ListCount < 18 And List1.ListCount <> 1 And List1.ListCount <> 0 Then
List1.Height = List1.Height - 360: List1.Top = List1.Top + 360
List2.Height = List2.Height - 360: List2.Top = List2.Top + 360
List3.Height = List3.Height - 360: List3.Top = List3.Top + 360
End If
If List1.ListCount > 1 Then
Label4.Caption = "TOTAL: " & Format(CCur(Mid(Label4, 10, Len(Label4)) - Mid(List3.Text, 3, Len(List3.Text))), "currency")
Else
Label4.Caption = Format(0, "currency")
End If
If List1.ListCount <> 0 Then List1.RemoveItem (List1.ListIndex)
If List2.ListCount <> 0 Then List2.RemoveItem (List2.ListIndex)
If List3.ListCount <> 0 Then List3.RemoveItem (List3.ListIndex)

If List1.ListCount <> 0 Then List1.ListIndex = List1.ListCount - 1
If List2.ListCount <> 0 Then List2.ListIndex = List2.ListCount - 1
If List3.ListCount <> 0 Then List3.ListIndex = List3.ListCount - 1
End Sub

Private Sub Form_Load()
var = 0
If tabprodutos.State = adStateOpen Then tabprodutos.Close
If tabmarca.State = adStateOpen Then tabmarca.Close
tabprodutos.Open "Produtos", conectar, adOpenKeyset, adLockOptimistic
tabmarca.Open "Marcas", conectar, adOpenKeyset, adLockOptimistic
End Sub

Private Sub Label11_Click()
Option1_Click
End Sub

Private Sub Label12_Click()
Option2_Click
End Sub

Private Sub Label13_Click()
Option3_Click
End Sub

Private Sub Label14_Click()
Call finalizar
End Sub

Private Sub Label5_Change()

If tabprodutos.State = adStateOpen Then tabprodutos.Close
If tabmarca.State = adStateOpen Then tabmarca.Close
codigo = Mid(Label5.Caption, 18, Len(Label5.Caption))
tabprodutos.Open "select * from Produtos where Codigo_barras = '" & codigo & "'"
If tabprodutos.RecordCount = 0 Then
            Label6.Caption = "QUANTIDADE: "
            lbl_subtotal.Caption = "SUBTOTAL: "
            Label7.Caption = ""
            Label8.Caption = "MARCA: "
            Label9.Caption = "PREÇO: "
            Label10.Caption = "VALIDADE: "
            Image1.Picture = LoadPicture(Empty)
Else
tabmarca.Open "select * from Marcas where Codigo =" & tabprodutos!Marca
Label7.Caption = "" & tabprodutos!nome
Label8.Caption = "MARCA: " & tabmarca!Marca
Label9.Caption = "PREÇO: " & Format(tabprodutos!Preco_unitario, "currency")
Label10.Caption = "VALIDADE: " & tabprodutos!Validade
Image1.Picture = LoadPicture(tabprodutos!Imagem)
End If

End Sub

Private Sub Label6_Change()
If Label6 = "QUANTIDADE:" Then Label6 = "QUANTIDADE: ": Exit Sub
If Label6 <> "QUANTIDADE: " Then
If Label9 <> "PREÇO: " Then
lbl_subtotal = "SUBTOTAL: " & Format(Mid(Label6.Caption, 12, Len(Label6.Caption)) * Mid(Label9.Caption, 7, Len(Label9)), "currency")
Else
lbl_subtotal = "SUBTOTAL: "
End If
End If
End Sub

Private Sub List1_Click()
If List1.ListCount = List2.ListCount And List1.ListCount = List3.ListCount Then
List2.ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex
End If
End Sub

Private Sub List2_Click()
If List2.ListCount = List1.ListCount And List2.ListCount = List3.ListCount Then
List1.ListIndex = List2.ListIndex
List3.ListIndex = List2.ListIndex
End If
End Sub

Private Sub List3_Click()
If List3.ListCount = List1.ListCount And List3.ListCount = List2.ListCount Then
List1.ListIndex = List3.ListIndex
List2.ListIndex = List3.ListIndex
End If
End Sub
Private Sub finalizar()

            Option1.value = False
            Option2.value = False
            Option3.value = False
            Option4.value = False
            List1.Clear
            List2.Clear
            List3.Clear
            Label5.Caption = "CÓDIGO DE BARRA: "
            Label6.Caption = "QUANTIDADE: "
            lbl_subtotal.Caption = "SUBTOTAL: "
            Label4.Caption = "TOTAL: R$ 0,00"
            Label8.Caption = "MARCA: "
            Label9.Caption = "PREÇO: "
            Label10.Caption = "VALIDADE: "
            Image1.Picture = LoadPicture(Empty)
            List1.Height = 390
            List2.Height = 390
            List3.Height = 390
            List1.Top = 7800
            List2.Top = 7800
            List3.Top = 7800
            Label6.BackColor = &HFFFFFF
            Label5.BackColor = &H800000
            Label5.ForeColor = &HFFFFFF
            Label6.ForeColor = &H0&
            MsgBox "COMPRA REALIZADA COM SUCESSO", vbInformation, "TRICON SUPERMERCADOS LTDA."

End Sub

Private Sub Option1_Click()
Call finalizar
End Sub

Private Sub Option2_Click()
Call finalizar
End Sub

Private Sub Option3_Click()
Call finalizar
End Sub

Private Sub Option4_Click()
Call finalizar
End Sub
