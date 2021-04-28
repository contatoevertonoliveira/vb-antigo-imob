Attribute VB_Name = "Module3"
Public cnn2 As New ADODB.Connection
Public Rs4 As New ADODB.Recordset

Sub conectar2()
    cnn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Dados.MDB"
    Rs4.CursorLocation = adUseClient
    Rs4.Open "Select * From Clientes", cnn2, adOpenKeyset, adLockOptimistic, adCmdText
End Sub
