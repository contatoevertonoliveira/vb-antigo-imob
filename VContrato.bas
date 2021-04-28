Attribute VB_Name = "Module5"
Public Dados As New ADODB.Connection
Public Tabela As New ADODB.Recordset

Sub AbreBd()
    Dados.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\dados\vcontrato.mdb"
    Tabela.CursorLocation = adUseClient
    Tabela.Open "Select * From VContrato", Dados, adOpenKeyset, adLockOptimistic, adCmdText
End Sub
