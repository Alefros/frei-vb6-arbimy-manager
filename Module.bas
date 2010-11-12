Attribute VB_Name = "Module2"
Global conectar1 As New ADODB.Connection
Global media As New ADODB.Recordset
Global mediana As New ADODB.Recordset
Global moda As New ADODB.Recordset
Global percentil As New ADODB.Recordset
Global quartil As New ADODB.Recordset
Global amplitude As New ADODB.Recordset
Global dp As New ADODB.Recordset
Global cv As New ADODB.Recordset
Global TESTE As New ADODB.Recordset

Function abrir()
calc_est = "provider=microsoft.jet.oledb.4.0; data source = " & App.Path & "\bd2.mdb"
            If conectar1.State = adStateOpen Then conectar1.Close
            conectar1.Open calc_est
End Function

