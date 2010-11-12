Attribute VB_Name = "Module1"
Global status, usuario As String
Global caminho As String
Global capsula As String
Global tabfrete As New ADODB.Recordset
Global tabuf As New ADODB.Recordset
Global tabpagamento As New ADODB.Recordset
Global tabmarca As New ADODB.Recordset
Global tabsegmento As New ADODB.Recordset
Global tabprodutos As New ADODB.Recordset
Global tabpedidodecompra As New ADODB.Recordset
Global tabsegmentos As New ADODB.Recordset
Global tabcid As New ADODB.Recordset
Global tabbairro As New ADODB.Recordset
Global tablocalizacao As New ADODB.Recordset
Global conectar As New ADODB.Connection
Global tabfornecedores As New ADODB.Recordset
Function abrir_banco()
            If conectar.State = adStateOpen Then conectar.Close
            capsula = "Provider=microsoft.jet.oledb.4.0;data source="
            caminho = capsula + App.Path & "\Supermercados.mdb"
            conectar.Open (caminho)
End Function

Function box()
            MsgBox "Informações " & status & " com sucesso", vbExclamation, "Tricon Supermercados LTDA."
End Function
Function voltar_botao()

           ' MDIForm1.Tbl_padrao.Buttons.Item(1).value = tbrUnpressed
           ' MDIForm1.Tbl_padrao.Buttons.Item(1).Caption = "Clientes"
          '  MDIForm1.Tbl_padrao.Buttons.Item(1).Style = tbrDropdown
          '  MDIForm1.Tbl_padrao.Buttons.Item(1).Image = 6
End Function
