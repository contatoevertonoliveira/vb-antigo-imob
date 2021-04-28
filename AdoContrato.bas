Attribute VB_Name = "Module1"
Public Cn As New ADODB.Connection
Public Rs As New ADODB.Recordset
Public Rs2 As New ADODB.Recordset
Public Rs3 As New ADODB.Recordset

Sub Conecta()
    Cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & "\\Maq5\c\Meus documentos\Documentos Backup\Programa Imobiliária\Dados\bdimobiliaria.mdb"
    Rs.CursorLocation = adUseClient
    Rs.Open "Select Nome From Prop where Nome", Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Rs2.CursorLocation = adUseClient
    Rs2.Open "Select Nome From Loc where Nome", Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Rs3.CursorLocation = adUseClient
    Rs3.Open "Select Locador From Contrato where Locador", Cn, adOpenKeyset, adLockOptimistic, adCmdText
End Sub
