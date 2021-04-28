Attribute VB_Name = "Module1"
Public Cn As New ADODB.Connection
Public Rs As New ADODB.Recordset
Public Rs2 As New ADODB.Recordset
Public Rs3 As New ADODB.Recordset
Public Rs5 As New ADODB.Recordset

Sub Conecta()
    Cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =" & App.Path & "\Dados\bdimobiliaria.mdb"
    Rs3.CursorLocation = adUseClient
    Rs3.Open "Select Locador From Contrato where Locador", Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Rs4.CursorLocation = adUseClient
    Rs4.Open "Select Vendedor From VContrato where Vendedor", Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Rs5.CursorLocation = adUseClient
    Rs5.Open "Select Codigo From Vencimentos where Codigo", Cn, adOpenKeyset, adLockOptimistic, adCmdText
End Sub
