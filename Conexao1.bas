Attribute VB_Name = "Module4"
Public Bd As New ADODB.Connection
Public Tb As New ADODB.Recordset

Sub conectar()
    Bd.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Pessoas.mdb"
    Tb.CursorLocation = adUseClient
    Tb.Open "Select * From Clientes", Bd, adOpenKeyset, adLockOptimistic, adCmdText
End Sub

