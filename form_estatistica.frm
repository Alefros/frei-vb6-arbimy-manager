VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form form_estatistica 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CALCULADORA ESTATÍSTICA"
   ClientHeight    =   9045
   ClientLeft      =   5475
   ClientTop       =   3465
   ClientWidth     =   10005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "COMPARAÇÃO"
      Height          =   615
      Left            =   3240
      TabIndex        =   27
      Top             =   2640
      Width           =   3735
      Begin VB.ComboBox cbo_com 
         Height          =   315
         ItemData        =   "form_estatistica.frx":0000
         Left            =   2520
         List            =   "form_estatistica.frx":0002
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cbo_de 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "form_estatistica.frx":0004
         Left            =   480
         List            =   "form_estatistica.frx":0006
         TabIndex        =   30
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "COM"
         Height          =   255
         Left            =   2040
         TabIndex        =   29
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "DE"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "GRÁFICO DE:"
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   2640
      Width           =   3015
      Begin VB.CommandButton cmd_limpar 
         Caption         =   "LIMPAR"
         Height          =   255
         Left            =   2040
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton opt_b 
         Caption         =   "BARRA"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton opt_s 
         Caption         =   "SETOR"
         Height          =   255
         Left            =   1080
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.ProgressBar pb_progresso 
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   8640
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   108
      Scrolling       =   1
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "OK"
      Height          =   255
      Left            =   9480
      TabIndex        =   15
      Top             =   8280
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   9855
      Begin VB.CommandButton Command1 
         Caption         =   "IMPORTAR"
         Height          =   255
         Left            =   8280
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cbo_setor 
         Height          =   315
         ItemData        =   "form_estatistica.frx":0008
         Left            =   4920
         List            =   "form_estatistica.frx":000F
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
      Begin VB.ComboBox cbo_ano 
         Height          =   315
         ItemData        =   "form_estatistica.frx":001A
         Left            =   840
         List            =   "form_estatistica.frx":001C
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "SETOR FINACEIRO"
         Height          =   255
         Left            =   3360
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "ANO"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.TextBox txt_percentil 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   7080
      TabIndex        =   14
      Top             =   8280
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid mfg_valores 
      Height          =   7215
      Left            =   7080
      TabIndex        =   16
      Top             =   960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   12726
      _Version        =   393216
      FormatString    =   "MES                         |    VALOR        "
   End
   Begin VB.Frame Frame1 
      Caption         =   "COMANDOS"
      Height          =   1695
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   6855
      Begin VB.CommandButton cmd_amplitude 
         Caption         =   "AMPLITUDE"
         Height          =   375
         Left            =   5160
         TabIndex        =   13
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmd_3_quartil 
         Caption         =   "3º QUARTIL"
         Height          =   375
         Left            =   3600
         TabIndex        =   12
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmd_1_quartil 
         Caption         =   "1º QUARTIL"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmd_2_quartil 
         Caption         =   "2º QUARTIL"
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmd_grafico_c 
         Caption         =   "GRÁFICO C."
         Height          =   375
         Left            =   5160
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmd_grafico 
         Caption         =   "GRÁFICO"
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmd_d_p 
         Caption         =   "D.P"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmd_c_v 
         Caption         =   "C.V"
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmd_percentil 
         Caption         =   "PERCENTIL"
         Height          =   375
         Left            =   5160
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_mediana 
         Caption         =   "MEDIANA"
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_media 
         Caption         =   "MEDIA"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_moda 
         Caption         =   "MODA"
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSChart20Lib.MSChart msc_grafico 
      Height          =   5175
      Left            =   120
      OleObjectBlob   =   "form_estatistica.frx":001E
      TabIndex        =   21
      Top             =   3360
      Width           =   6855
   End
End
Attribute VB_Name = "form_estatistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lista, itens, l_dp, i_dp, lista_m, ano, setor, v_moda, itens_m, cont, soma, adicionar, var, i, validar, pos, veri, tempo, resultado, repeticao, maior, menor, mensagem, codigo, cl, a, b, c, d, e, f, g, h, j, k As String

Private Sub cmd_grafico_c_Click()
            If cbo_com = "" Then MsgBox "SELECIONE UM ANO A SER COMPARADO", vbInformation, "TRICON SUPERMERCADOS LTDA.": Exit Sub
            cl = "GRAFICO"
            Call grafico_c
End Sub

Private Sub cmd_limpar_Click()
            If msc_grafico.ColumnCount <> 0 Then msc_grafico.ColumnCount = 0
            If msc_grafico.ColumnLabel <> "" Then msc_grafico.ColumnLabel = ""
            If msc_grafico.Footnote <> "" Then msc_grafico.Footnote = ""
            If msc_grafico.RowCount <> 1 Then msc_grafico.RowCount = 1
            If msc_grafico.RowLabel <> "" Then msc_grafico.RowLabel = ""
            If msc_grafico.Title <> "" Then msc_grafico.TitleText = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then cmd_ok_Click
End Sub

Private Sub Form_Load()
        Call abrir
        Call abrir_
        Call carregar_lista
        pb_progresso.Visible = True
        i = 2007
        ano = Right(Date$, 4)
        Do Until i = ano + 1
        cbo_ano.AddItem (i)
        cbo_com.AddItem (i)
        cbo_de.AddItem (i)
        i = i + 1
        Loop
        i = ""
        ano = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
            pb_progresso.value = 0
            pb_progresso.Visible = False
            Call TESTE_
End Sub

Private Sub flex()
            mfg_valores.Clear
            mfg_valores.FormatString = "MES                         |    VALOR        "
            i = 0
            Call abrir_
            mfg_valores.Rows = 2
            If cl = "MEDIA" Then media.MoveFirst
            If cl = "MEDIANA" Then mediana.MoveFirst
            If cl = "MODA" Then mediana.MoveFirst
            If cl = "GRAFICO" Then mediana.MoveFirst
            If cl = "PERCENTIL" Then percentil.MoveFirst
            If cl = "1º QUARTIL" Or cl = "2º QUARTIL" Or cl = "3º QUARTIL" Then quartil.MoveFirst
            If cl = "AMPLITUDE" Then amplitude.MoveFirst
            If cl = "DESVIO PADRÃO" Then dp.MoveFirst
            If cl = "CV" Then cv.MoveFirst
            i = 0
            
            Do Until media.EOF = True
            mfg_valores.TextMatrix(mfg_valores.Rows - 1, 0) = itens(i)
            If cl = "MEDIA" Then mfg_valores.TextMatrix(mfg_valores.Rows - 1, 1) = Format(media!num, "currency")
            If cl = "MEDIANA" Then mfg_valores.TextMatrix(mfg_valores.Rows - 1, 1) = Format(mediana!num, "currency")
            If cl = "MODA" Then mfg_valores.TextMatrix(mfg_valores.Rows - 1, 1) = Format(mediana!num, "currency")
            If cl = "GRAFICO" Then mfg_valores.TextMatrix(mfg_valores.Rows - 1, 1) = Format(mediana!num, "currency")
            If cl = "PERCENTIL" Then mfg_valores.TextMatrix(mfg_valores.Rows - 1, 1) = Format(percentil!num, "currency")
            If cl = "1º QUARTIL" Or cl = "2º QUARTIL" Or cl = "3º QUARTIL" Then mfg_valores.TextMatrix(mfg_valores.Rows - 1, 1) = Format(quartil!num, "currency")
            If cl = "AMPLITUDE" Then mfg_valores.TextMatrix(mfg_valores.Rows - 1, 1) = Format(amplitude!num, "currency")
            If cl = "DESVIO PADRÃO" Then mfg_valores.TextMatrix(mfg_valores.Rows - 1, 1) = Format(dp!num, "currency")
            If cl = "CV" Then mfg_valores.TextMatrix(mfg_valores.Rows - 1, 1) = Format(cv!num, "currency")
            mfg_valores.Rows = mfg_valores.Rows + 1
            i = i + 1
            
            media.MoveNext
            If cl = "MEDIANA" Then mediana.MoveNext
            If cl = "MODA" And moda.EOF = False Then mediana.MoveNext
            If cl = "GRAFICO" And moda.EOF = False Then mediana.MoveNext
            If cl = "PERCENTIL" Then percentil.MoveNext
            If cl = "1º QUARTIL" Or cl = "2º QUARTIL" Or cl = "3º QUARTIL" Then quartil.MoveNext
            If cl = "AMPLITUDE" Then amplitude.MoveNext
            If cl = "DESVIO PADRÃO" Then dp.MoveNext
            If cl = "CV" Then cv.MoveNext
            Loop
            If cl = "GRAFICO" Then Exit Sub
            If cl <> "MODA" Then Call flex_adicional
            If cl = "MODA" Then Call flex_moda
            
End Sub

Private Sub flex_moda()

                mfg_valores.TextMatrix(mfg_valores.Rows - 1, 0) = ""
                mfg_valores.TextMatrix(mfg_valores.Rows - 1, 1) = ""
                cont = 1
                Do Until cont = lista_m.Count
                If v_moda <> "NÃO" Then
                mfg_valores.Rows = mfg_valores.Rows + 1
                mfg_valores.TextMatrix(mfg_valores.Rows - 1, 0) = cl
                mfg_valores.TextMatrix(mfg_valores.Rows - 1, 1) = Format(itens_m(cont), "currency")
                End If
                cont = cont + 1
                Loop
                mfg_valores.Rows = mfg_valores.Rows + 1
                If v_moda <> "NÃO" Then mfg_valores.TextMatrix(mfg_valores.Rows - 1, 0) = "REPETIÇÕES"
                If v_moda <> "NÃO" Then mfg_valores.TextMatrix(mfg_valores.Rows - 1, 1) = codigo
                If v_moda <> "NÃO" Then mfg_valores.Rows = mfg_valores.Rows + 1
                mfg_valores.TextMatrix(mfg_valores.Rows - 1, 0) = cl
                mfg_valores.TextMatrix(mfg_valores.Rows - 1, 1) = veri
                
End Sub


Private Sub carregar_lista()
            lista = ""
            itens = ""
            Set lista = CreateObject("scripting.Dictionary")
            lista.Add "1", "JANEIRO"
            lista.Add "2", "FEVEREIRO"
            lista.Add "3", "MARÇO"
            lista.Add "4", "ABRIL"
            lista.Add "5", "MAIO"
            lista.Add "6", "JUNHO"
            lista.Add "7", "JULHO"
            lista.Add "8", "AGOSTO"
            lista.Add "9", "SETEMBRO"
            lista.Add "10", "OUTUBRO"
            lista.Add "11", "NOVEMBRO"
            lista.Add "12", "DEZEMBRO"
            itens = lista.Items
End Sub

Private Sub opt_b_Click()
            If msc_grafico.ColumnCount <> 0 Then
            msc_grafico.chartType = VtChChartType2dBar
            End If
End Sub

Private Sub opt_s_Click()
            If msc_grafico.ColumnCount <> 0 Then
            msc_grafico.chartType = VtChChartType2dPie
            End If
End Sub

Private Sub cmd_1_quartil_Click()
            cl = "1º QUARTIL"
            Call calc_quartil
            If validar = "SIM" Then Call flex
End Sub

Private Sub cmd_2_quartil_Click()
            cl = "2º QUARTIL"
            Call calc_quartil
            If validar = "SIM" Then Call flex
End Sub

Private Sub cmd_3_quartil_Click()
            cl = "3º QUARTIL"
            Call calc_quartil
            If validar = "SIM" Then Call flex
End Sub

Private Sub cmd_amplitude_Click()
            cl = "AMPLITUDE"
            Call calc_amplitude
            If validar = "SIM" Then Call flex
End Sub

Private Sub cmd_media_Click()
            cl = "MEDIA"
            Call calc_media
            If validar = "SIM" Then Call flex
End Sub

Private Sub cmd_d_p_Click()
            cl = "DESVIO PADRÃO"
            Call calc_d_p
            If validar = "SIM" Then Call flex
End Sub

Private Sub cmd_mediana_Click()
            cl = "MEDIANA"
            Call calc_mediana
            If validar = "SIM" Then Call flex
End Sub

Private Sub cmd_c_v_Click()
        cl = "CV"
        Call calc_cv
        If validar = "SIM" Then Call flex
End Sub

Private Sub cmd_moda_Click()
            cl = "MODA"
            Call calc_moda
            If validar = "SIM" Then Call flex
End Sub

Private Sub cmd_grafico_Click()
            cl = "GRAFICO"
            Call grafico_
End Sub

Private Sub cmd_ok_Click()
            If txt_percentil = "" Then txt_percentil.SetFocus: Exit Sub
            If var = 1111 Then
            cl = "PERCENTIL"
            If txt_percentil > 100 Then txt_percentil = 100
            If txt_percentil < 1 Then txt_percentil = 1
            Call calc_percentil
            If validar = "SIM" Then Call flex
            cmd_percentil.SetFocus
            End If
End Sub

Private Sub cmd_percentil_Click()
            var = 1111
            txt_percentil.SetFocus
End Sub

Private Sub Command1_Click()
            If cbo_ano.Text <> "" And cbo_setor.Text <> "" Then
            pb_progresso.value = 0
            Call conferir
            If resultado <> "ACEITO" Then MsgBox "NENHUM VALOR EXISTENTE": Exit Sub
            cbo_de.Text = cbo_ano.Text
            ano = cbo_ano.Text
            setor = cbo_setor.Text
            Call TESTE_
            Call t_anos
            
            Else
            MsgBox "SELECIONE UM ANO E UM SETOR FINANCEIRO"
            End If
End Sub

      '/////////////////////////////MEDIA//////////////////////////////////////'


Private Sub calc_media()
            cont = 1
            soma = 0
            var = ""
            resultado = ""
            pos = ""
            If media.State = adStateOpen Then media.Close
            media.Open "media", conectar1, adOpenKeyset, adLockOptimistic
            Do Until media.EOF = True
            If media!num <> "" Then
            soma = CCur(soma) + media!num
            End If
            media.MoveNext
            cont = cont + 1
            Loop
            If cont = 1 Then MsgBox "DADOS NÃO EXISTENTE, IMPORTE-OS", vbInformation, "TRICON SUPERMERCADOS LTDA.": validar = "NÃO": Exit Sub
            cont = cont - 1
            var = soma / cont
            i = 1
            pos = InStr(i, var, ",", vbTextCompare)
            If pos <> 0 Then
            If Len(var) > 5 Then resultado = Left(var, pos + 2)
            End If
            
            If resultado = "" Then resultado = var
            validar = "SIM"
            cont = 1
            soma = 0
            var = ""
End Sub


      '/////////////////////////////MEDIANA//////////////////////////////////////'


Private Sub calc_mediana()

            If mediana.State = adStateOpen Then mediana.Close
            mediana.Open "select * from mediana order by num"
            Set lista = CreateObject("scripting.Dictionary")
            
            b = 0
            Do Until mediana.EOF = True
            lista.Add "" & b & "", "" & mediana!num & ""
            b = b + 1
            mediana.MoveNext
            Loop
            '''''''''AQUI'''''''
            If b = 0 Then MsgBox "DADOS NÃO EXISTENTE, IMPORTE-OS", vbInformation, "TRICON SUPERMERCADOS LTDA.": validar = "NÃO": Exit Sub
            cont = b / 2
            
            j = lista.Items
                i = 1
                pos = InStr(i, cont, ",", vbTextCompare)
                If pos <> 0 Then
                resultado = j((CCur(cont) + "0,5") - 1)
                Exit Sub
                End If
                
                k = j(cont - 1)
                L = j(CCur(cont - 1) + 1)
                resultado = (CCur(k) + CCur(L)) / 2
                resultado = Format(resultado, "currency")
                cont = ""
                validar = "SIM"
APAGAR:
            j = ""
            k = ""
            a = ""
            b = ""

End Sub

      '/////////////////////////////MODA//////////////////////////////////////'

Private Sub calc_moda()
            var = 0
            cont = 0
            Set lista_m = CreateObject("scripting.Dictionary")
            repeticao = 0
            If moda.State = adStateOpen Then moda.Close
            moda.Open "select * from moda order by repeticao"
            If moda!num = "" Then MsgBox "DADOS NÃO EXISTENTE, IMPORTE-OS", vbInformation, "TRICON SUPERMERCADOS LTDA.": validar = "NÃO": Exit Sub
            moda.MoveLast
            codigo = moda!repeticao: lista_m.Add "" & var & "", "" & moda!num & ""
            moda.MoveFirst
            var = var + 1
            cont = cont + 1
            Do Until moda.EOF = True
            If codigo = moda!repeticao Then
            lista_m.Add "" & var & "", "" & moda!num & "": var = var + 1
            End If
            cont = cont + 1
            moda.MoveNext
            Loop
            
            If var = 1 Then veri = "AMODAL": v_moda = "NÃO": validar = "SIM": Exit Sub
            If var = cont Then veri = "AMODAL": v_moda = "NÃO": validar = "SIM": Exit Sub
            itens_m = lista_m.Items
            
            If lista_m.Count - 1 = 0 Then veri = "AMODAL"
            If lista_m.Count - 1 = 1 Then veri = "UNIMODAL"
            If lista_m.Count - 1 = 2 Then veri = "BIMODAL"
            If lista_m.Count - 1 = 3 Then veri = "TRIMODAL"
            If lista_m.Count - 1 >= 4 Then veri = "POLIMODAL"
            validar = "SIM"

End Sub


'/////////////////////////////////PERCENTIL//////////////////////////////////'

Private Sub calc_percentil()
            
            Set lista_m = CreateObject("Scripting.Dictionary")
            If percentil.State = adStateOpen Then percentil.Close
            percentil.Open "select * from percentil order by num"
            cont = 1
            Do Until percentil.EOF = True
            lista_m.Add "" & cont & "", "" & percentil!num & ""
            cont = cont + 1
            percentil.MoveNext
            Loop
            If cont = 1 Then MsgBox "DADOS NÃO EXISTENTE, IMPORTE-OS", vbInformation, "TRICON SUPERMERCADOS LTDA.": validar = "NÃO": Exit Sub
            cont = cont - 1
            total = (txt_percentil / 100) * cont
            itens_m = lista_m.Items
            percentil.MoveFirst
            validar = "SIM"
            i = 1
            pos = InStr(i, total, ",", vbTextCompare)
            If pos <> 0 Then
            resultado = itens_m(Left(total, pos - 1))
            Exit Sub
            End If
            resultado = itens_m(total - 1)
End Sub


'/////////////////////////////////QUARTIL//////////////////////////////////'

Private Sub calc_quartil()

            Set lista_m = CreateObject("Scripting.Dictionary")
            If quartil.State = adStateOpen Then quartil.Close
            quartil.Open "select * from quartil order by num"
            cont = 0
            Do Until quartil.EOF = True
            lista_m.Add "" & cont & "", "" & quartil!num & ""
            cont = cont + 1
            quartil.MoveNext
            Loop
            If cont = 0 Then MsgBox "DADOS NÃO EXISTENTE, IMPORTE-OS", vbInformation, "TRICON SUPERMERCADOS LTDA.": validar = "NÃO": Exit Sub
            If cont = 1 Then MsgBox "DADOS INSULFICIENTE", vbInformation, "TRICON SUPERMERCADOS LTDA.": validar = "NÃO": Exit Sub
            cont = cont - 1
            itens_m = lista_m.Items
            resultado = (Left(cl, 1) / 4) * cont
            i = 1
            pos = InStr(i, resultado, ",", vbTextCompare)
            If pos <> 0 Then
            resultado = itens_m(Left(resultado, pos - 1))
            Else
            resultado = (CCur(itens_m(resultado - 1)) + CCur(itens_m(resultado))) / 2
            End If
            validar = "SIM"
            resultado = Format(resultado, "currency")
End Sub

'/////////////////////////////////QUARTIL//////////////////////////////////'

Private Sub calc_amplitude()
            If amplitude.State = adStateOpen Then amplitude.Close
            amplitude.Open "select * from amplitude order by num"
            If amplitude!num = "" Then MsgBox "DADOS NÃO EXISTENTE, IMPORTE-OS", vbInformation, "TRICON SUPERMERCADOS LTDA.": validar = "NÃO": Exit Sub
            menor = amplitude!num
            amplitude.MoveLast
            maior = amplitude!num
            resultado = maior - menor
            validar = "SIM"
            resultado = Format(resultado, "currency")
End Sub

'/////////////////////////////////DESVIO PADRÃO//////////////////////////////////'

Private Sub calc_d_p()
            Call calc_media
            If validar <> "SIM" Then Exit Sub
            Set lista_m = CreateObject("scripting.Dictionary")
            Set l_dp = CreateObject("scripting.Dictionary")
            
            If dp.State = adStateOpen Then dp.Close
            dp.Open "Desvio_Padrao", conectar1, adOpenKeyset, adLockOptimistic
            var = 0
            cont = 0
            soma = 0
            Do Until dp.EOF = True
            l_dp.Add "" & cont & "", "" & dp!num - resultado & ""
            cont = cont + 1
            dp.MoveNext
            Loop
            dp.MoveFirst: If cont = 0 Then MsgBox "DADOS NÃO EXISTENTE, IMPORTE-OS", vbInformation, "TRICON SUPERMERCADOS LTDA.": validar = "NÃO": Exit Sub
            cont = 0
            i_dp = l_dp.Items
            Do Until cont = l_dp.Count
            lista_m.Add "" & cont & "", "" & i_dp(cont) * i_dp(cont) & ""
            cont = cont + 1
            Loop
            cont = 0
            itens_m = lista_m.Items
            Do Until cont = l_dp.Count
            soma = CCur(soma + itens_m(cont))
            cont = cont + 1
            Loop
            soma = soma / cont
            i = 1
            pos = InStr(i, soma, ",", vbTextCompare)
            If pos <> 0 Then
            soma = Left(soma, pos + 2)
            End If
            resultado = Sqr(soma)
            pos = InStr(i, resultado, ",", vbTextCompare)
            If pos <> 0 Then
            resultado = Left(resultado, pos + 2)
            End If
            validar = "SIM"
            resultado = Format(resultado, "currency")
End Sub

'/////////////////////////////////CV//////////////////////////////////'

Private Sub calc_cv()
            Call calc_d_p
            If resultado = "" Then validar = "NÃO": Exit Sub
            j = resultado
            Call calc_media
            k = resultado
            resultado = j / k * 100
            i = 1
            pos = InStr(i, resultado, ",", vbTextCompare)
            If pos <> 0 Then
            resultado = Left(resultado, pos + 2)
            End If
            resultado = resultado & "%"
            validar = "SIM"
End Sub

'/////////////////////////////////GRÁFICO//////////////////////////////////'

Private Sub grafico_()

            If media.State = adStateOpen Then media.Close
            media.Open "media", conectar1, adOpenKeyset, adLockOptimistic
            If media!num = "" Then MsgBox "DADOS NÃO EXISTENTE, IMPORTE-OS", vbInformation, "TRICON SUPERMERCADOS LTDA.": validar = "NÃO": Exit Sub
            If opt_b.value = True Then msc_grafico.chartType = VtChChartType2dBar
            If opt_s.value = True Then msc_grafico.chartType = VtChChartType2dPie
            msc_grafico.TitleText = setor & "  -  " & ano
            msc_grafico.Footnote = "FONTE: TRICON SUPERMERCADOS LTDA."
            
            i = 1
            Set lista_m = CreateObject("scripting.Dictionary")
            TESTE.MoveFirst
            Do Until TESTE.EOF = True
            If cbo_de = 2007 Then adicionar = TESTE!t_2007
            If cbo_de = 2008 Then adicionar = TESTE!t_2008
            If cbo_de = 2009 Then adicionar = TESTE!t_2009
            If cbo_de = 2010 Then adicionar = TESTE!t_2010
            If cbo_de = 2011 Then adicionar = TESTE!t_2011
            If cbo_de = 2012 Then adicionar = TESTE!t_2012
            If cbo_de = 2013 Then adicionar = TESTE!t_2013
            If cbo_de = 2014 Then adicionar = TESTE!t_2014
            If cbo_de = 2015 Then adicionar = TESTE!t_2015
            If cbo_de = 2016 Then adicionar = TESTE!t_2016
            If cbo_de = 2017 Then adicionar = TESTE!t_2017
            If cbo_de = 2018 Then adicionar = TESTE!t_2018
            If cbo_de = 2019 Then adicionar = TESTE!t_2019
            If cbo_de = 2020 Then adicionar = TESTE!t_2020
            
            If Len(adicionar) <> 0 Then lista_m.Add "" & i & "", "" & adicionar & ""
            TESTE.MoveNext
            i = i + 1
            Loop
            
            i = 0
            
            msc_grafico.RowCount = 1
            msc_grafico.ColumnCount = 1
            msc_grafico.Column = 1
            
            media.MoveFirst
            Do Until i = lista_m.Count
                msc_grafico.ColumnLabel = itens(i)
                msc_grafico.Data = media!num
            media.MoveNext
            msc_grafico.ColumnCount = msc_grafico.ColumnCount + 1
            msc_grafico.Column = msc_grafico.Column + 1
            i = i + 1
            Loop
            msc_grafico.ColumnCount = msc_grafico.ColumnCount - 1
            validar = "SIM"
            Call flex
End Sub

'//////////////////////////////GRÁFICO COMPARATIVO///////////////////////////////'

Private Sub grafico_c()
            Call calc_media
            codigo = ""
            var = 0
            If validar = "NÃO" Then Exit Sub
            msc_grafico.Footnote = "FONTE: TRICON SUPERMERCADOS LTDA."
            msc_grafico.ColumnCount = 2
            msc_grafico.RowCount = 1
            msc_grafico.Column = 1
            If opt_b.value = True Then msc_grafico.chartType = VtChChartType2dBar
            If opt_s.value = True Then msc_grafico.chartType = VtChChartType2dPie
VOLTA:
            i = 0
            var = var + 1
            
            
            If media.State = adStateOpen Then media.Close
            media.Open "media", conectar1, adOpenKeyset, adLockOptimistic
            If media!num = "" Then MsgBox "DADOS NÃO EXISTENTE, IMPORTE-OS", vbInformation, "TRICON SUPERMERCADOS LTDA.": validar = "NÃO": Exit Sub
            msc_grafico.TitleText = "COMPARAÇÃO ENTRE " & cbo_de & " E " & cbo_com & " "
            
            If var = 1 Then msc_grafico.ColumnLabel = cbo_de
            If var = 2 Then msc_grafico.ColumnLabel = cbo_com
            msc_grafico.Data = resultado
            If msc_grafico.Column <> 2 Then msc_grafico.Column = msc_grafico.Column + 1
            i = i + 1
            If codigo = "" Then codigo = cbo_ano.Text
            If var = 1 Then cbo_ano.Text = cbo_com.Text: pb_progresso.value = 0: Call TESTE_: t_anos: g_resultado
            If var = 1 Then GoTo VOLTA:
            cbo_ano = codigo
            validar = "SIM"
            cbo_ano.Text = cbo_de.Text: pb_progresso.value = 0: pb_progresso.Visible = False: Call TESTE_: t_anos: g_resultado
            pb_progresso.Visible = True
            Call flex
            i = ""
            var = ""
End Sub
Private Sub g_resultado()
            i = 0
            resultado = ""
            If media.State = adStateOpen Then media.Close
            media.Open "media", conectar1, adOpenKeyset, adLockOptimistic
            Do Until media.EOF
            resultado = CCur(resultado + media!num)
            i = i + 1
            media.MoveNext
            Loop
            resultado = resultado / i
            i = 1
            pos = InStr(i, resultado, ",", vbTextCompare)
            If pos <> 0 Then
            resultado = Left(resultado, pos + 2)
            End If
            i = 0

End Sub



Private Sub flex_adicional()
                mfg_valores.TextMatrix(mfg_valores.Rows - 1, 0) = ""
                mfg_valores.TextMatrix(mfg_valores.Rows - 1, 1) = ""
                mfg_valores.Rows = mfg_valores.Rows + 1
                mfg_valores.TextMatrix(mfg_valores.Rows - 1, 0) = cl
                If cl <> "CV" Then mfg_valores.TextMatrix(mfg_valores.Rows - 1, 1) = Format(resultado, "currency")
                If cl = "CV" Then mfg_valores.TextMatrix(mfg_valores.Rows - 1, 1) = resultado
End Sub

Private Sub abrir_()
            If media.State = adStateOpen Then media.Close
            media.Open "media", conectar1, adOpenKeyset, adLockOptimistic
            
            If mediana.State = adStateOpen Then mediana.Close
            mediana.Open "mediana", conectar1, adOpenKeyset, adLockOptimistic

            If moda.State = adStateOpen Then moda.Close
            moda.Open "moda", conectar1, adOpenKeyset, adLockOptimistic
            
            If percentil.State = adStateOpen Then percentil.Close
            percentil.Open "percentil", conectar1, adOpenKeyset, adLockOptimistic
            
            If quartil.State = adStateOpen Then quartil.Close
            quartil.Open "quartil", conectar1, adOpenKeyset, adLockOptimistic
            
            If amplitude.State = adStateOpen Then amplitude.Close
            amplitude.Open "amplitude", conectar1, adOpenKeyset, adLockOptimistic
            
            If dp.State = adStateOpen Then dp.Close
            dp.Open "Desvio_Padrao", conectar1, adOpenKeyset, adLockOptimistic
            
            If cv.State = adStateOpen Then cv.Close
            cv.Open "CV", conectar1, adOpenKeyset, adLockOptimistic
            
            If TESTE.State = adStateOpen Then TESTE.Close
            TESTE.Open "TESTE", conectar1, adOpenKeyset, adLockOptimistic
End Sub

Private Sub TESTE_()
            If media.State = adStateOpen Then media.Close:
            If mediana.State = adStateOpen Then mediana.Close:
            If moda.State = adStateOpen Then moda.Close:
            If percentil.State = adStateOpen Then percentil.Close:
            If quartil.State = adStateOpen Then quartil.Close:
            If amplitude.State = adStateOpen Then amplitude.Close:
            If dp.State = adStateOpen Then dp.Close:
            If cv.State = adStateOpen Then cv.Close:
            media.Open "delete * from media":                       pb_progresso.value = pb_progresso.value + 1
            mediana.Open "delete * from mediana":                   pb_progresso.value = pb_progresso.value + 1
            moda.Open "delete * from moda":                         pb_progresso.value = pb_progresso.value + 1
            percentil.Open "delete * from percentil":               pb_progresso.value = pb_progresso.value + 1
            quartil.Open "delete * from quartil":                   pb_progresso.value = pb_progresso.value + 1
            amplitude.Open "delete * from amplitude":               pb_progresso.value = pb_progresso.value + 1
            dp.Open "delete * from Desvio_Padrao":                  pb_progresso.value = pb_progresso.value + 1
            cv.Open "delete * from CV":                             pb_progresso.value = pb_progresso.value + 1
End Sub

Private Sub t_anos()
            If cbo_setor = "TESTE" Then TESTE.MoveFirst
            
            Do Until TESTE.EOF = True
            
            If cbo_ano = 2007 Then adicionar = TESTE!t_2007
            If cbo_ano = 2008 Then adicionar = TESTE!t_2008
            If cbo_ano = 2009 Then adicionar = TESTE!t_2009
            If cbo_ano = 2010 Then adicionar = TESTE!t_2010
            If cbo_ano = 2011 Then adicionar = TESTE!t_2011
            If cbo_ano = 2012 Then adicionar = TESTE!t_2012
            If cbo_ano = 2013 Then adicionar = TESTE!t_2013
            If cbo_ano = 2014 Then adicionar = TESTE!t_2014
            If cbo_ano = 2015 Then adicionar = TESTE!t_2015
            If cbo_ano = 2016 Then adicionar = TESTE!t_2016
            If cbo_ano = 2017 Then adicionar = TESTE!t_2017
            If cbo_ano = 2018 Then adicionar = TESTE!t_2018
            If cbo_ano = 2019 Then adicionar = TESTE!t_2019
            If cbo_ano = 2020 Then adicionar = TESTE!t_2020
            
            If adicionar <> "" Then conectar1.Execute "insert into media(num)            values ('" & adicionar & "')":              pb_progresso.value = pb_progresso.value + 1
            If adicionar <> "" Then conectar1.Execute "insert into mediana(num)          values ('" & adicionar & "')":              pb_progresso.value = pb_progresso.value + 1
            If adicionar <> "" Then conectar1.Execute "insert into quartil(num)          values ('" & adicionar & "')":              pb_progresso.value = pb_progresso.value + 1
            If adicionar <> "" Then conectar1.Execute "insert into percentil(num)        values ('" & adicionar & "')":              pb_progresso.value = pb_progresso.value + 1
            If adicionar <> "" Then conectar1.Execute "insert into amplitude(num)        values ('" & adicionar & "')":              pb_progresso.value = pb_progresso.value + 1
            If adicionar <> "" Then conectar1.Execute "insert into Desvio_Padrao(num)    values ('" & adicionar & "')":              pb_progresso.value = pb_progresso.value + 1
            If adicionar <> "" Then conectar1.Execute "insert into CV(num)               values ('" & adicionar & "')":              pb_progresso.value = pb_progresso.value + 1

            If moda.State = adStateOpen Then moda.Close
            moda.Open "select * from moda where num = '" & adicionar & "'"
            If moda.RecordCount <> 0 Then
            moda!repeticao = moda!repeticao + 1
            moda.Update
            Else
            If adicionar <> "" Then conectar1.Execute "insert into moda(num) values ('" & adicionar & "')"
            End If
            

            pb_progresso.value = pb_progresso.value + 1
            If cbo_setor = "TESTE" Then TESTE.MoveNext
            Loop
            pb_progresso.value = pb_progresso.max
            
End Sub

Private Sub conferir()
            If TESTE.State = adStateOpen Then TESTE.Close
            TESTE.Open "TESTE", conectar1, adOpenKeyset, adLockOptimistic
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If cbo_ano = 2007 Then
            i = 1
            cont = 1
            Do Until i = 13
            If Len(TESTE!t_2007) <> 0 Then resultado = "ACEITO": Exit Sub
            If cbo_setor = "TESTE" Then TESTE.MoveNext
            i = i + 1
            Loop
            resultado = "RECUSADO": Exit Sub
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If cbo_ano = 2008 Then
            i = 1
            cont = 1
            Do Until i = 13
            If Len(TESTE!t_2008) <> 0 Then resultado = "ACEITO": Exit Sub
            If cbo_setor = "TESTE" Then TESTE.MoveNext
            i = i + 1
            Loop
            resultado = "RECUSADO": Exit Sub
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If cbo_ano = 2009 Then
            i = 1
            cont = 1
            Do Until i = 13
            If Len(TESTE!t_2009) <> 0 Then resultado = "ACEITO": Exit Sub
            If cbo_setor = "TESTE" Then TESTE.MoveNext
            i = i + 1
            Loop
            resultado = "RECUSADO": Exit Sub
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If cbo_ano = 2010 Then
            i = 1
            cont = 1
            Do Until i = 13
            If Len(TESTE!t_2010) <> 0 Then resultado = "ACEITO": Exit Sub
            If cbo_setor = "TESTE" Then TESTE.MoveNext
            i = i + 1
            Loop
            resultado = "RECUSADO": Exit Sub
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If cbo_ano = 2011 Then
            i = 1
            cont = 1
            Do Until i = 13
            If Len(TESTE!t_2011) <> 0 Then resultado = "ACEITO": Exit Sub
            If cbo_setor = "TESTE" Then TESTE.MoveNext
            i = i + 1
            Loop
            resultado = "RECUSADO": Exit Sub
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If cbo_ano = 2012 Then
            i = 1
            cont = 1
            Do Until i = 13
            If Len(TESTE!t_2012) <> 0 Then resultado = "ACEITO": Exit Sub
            If cbo_setor = "TESTE" Then TESTE.MoveNext
            i = i + 1
            Loop
            resultado = "RECUSADO": Exit Sub
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If cbo_ano = 2013 Then
            i = 1
            cont = 1
            Do Until i = 13
            If Len(TESTE!t_2013) <> 0 Then resultado = "ACEITO": Exit Sub
            If cbo_setor = "TESTE" Then TESTE.MoveNext
            i = i + 1
            Loop
            resultado = "RECUSADO": Exit Sub
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If cbo_ano = 2014 Then
            i = 1
            cont = 1
            Do Until i = 13
            If Len(TESTE!t_2014) <> 0 Then resultado = "ACEITO": Exit Sub
            If cbo_setor = "TESTE" Then TESTE.MoveNext
            i = i + 1
            Loop
            resultado = "RECUSADO": Exit Sub
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If cbo_ano = 2015 Then
            i = 1
            cont = 1
            Do Until i = 13
            If Len(TESTE!t_2015) <> 0 Then resultado = "ACEITO": Exit Sub
            If cbo_setor = "TESTE" Then TESTE.MoveNext
            i = i + 1
            Loop
            resultado = "RECUSADO": Exit Sub
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If cbo_ano = 2016 Then
            i = 1
            cont = 1
            Do Until i = 13
            If Len(TESTE!t_2016) <> 0 Then resultado = "ACEITO": Exit Sub
            If cbo_setor = "TESTE" Then TESTE.MoveNext
            i = i + 1
            Loop
            resultado = "RECUSADO": Exit Sub
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If cbo_ano = 2017 Then
            i = 1
            cont = 1
            Do Until i = 13
            If Len(TESTE!t_2017) <> 0 Then resultado = "ACEITO": Exit Sub
            If cbo_setor = "TESTE" Then TESTE.MoveNext
            i = i + 1
            Loop
            resultado = "RECUSADO": Exit Sub
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If cbo_ano = 2018 Then
            i = 1
            cont = 1
            Do Until i = 13
            If Len(TESTE!t_2018) <> 0 Then resultado = "ACEITO": Exit Sub
            If cbo_setor = "TESTE" Then TESTE.MoveNext
            i = i + 1
            Loop
            resultado = "RECUSADO": Exit Sub
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If cbo_ano = 2019 Then
            i = 1
            cont = 1
            Do Until i = 13
            If Len(TESTE!t_2019) <> 0 Then resultado = "ACEITO": Exit Sub
            If cbo_setor = "TESTE" Then TESTE.MoveNext
            i = i + 1
            Loop
            resultado = "RECUSADO": Exit Sub
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If cbo_ano = 2020 Then
            i = 1
            cont = 1
            Do Until i = 13
            If Len(TESTE!t_2020) <> 0 Then resultado = "ACEITO": Exit Sub
            If cbo_setor = "TESTE" Then TESTE.MoveNext
            i = i + 1
            Loop
            resultado = "RECUSADO"
            End If
End Sub

Private Sub txt_percentil_KeyPress(KeyAscii As Integer)
            If KeyAscii < 48 Or KeyAscii > 57 Then
                If KeyAscii <> 8 And KeyAscii <> 44 Then
                    If KeyAscii <> 45 Then
                            KeyAscii = 0
                    End If
                End If
            End If
        
        If KeyAscii = 46 Then KeyAscii = 44
            If KeyAscii = 44 And InStr(txt_percentil.Text, ",") <> 0 Then
            KeyAscii = 0
            Exit Sub
            End If
            If KeyAscii = 45 Then
            If Len(txt_salario) >= 1 Then KeyAscii = 0
            End If
End Sub
