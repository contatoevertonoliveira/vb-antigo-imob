Attribute VB_Name = "Modrecordsets"
Public sCamin1 As String    'caminho ao banco de dados
Public vTip As String       'tipo de base de dados a ser carregada F(FIREBIRD) ou d(DTA-ACCESS)
Public vTrvTec As String    'define o tipo de travamento do teclado (ctrl + alt + del) a ser usado
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'conexao ADO
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public con As New ADODB.Connection    'conexao com o banco via ADO
Public Lc1 As ADODB.Recordset         'recordset com a maquina filtrada (LOGADA)
Public Lc2 As ADODB.Recordset         'recordset com os clientes
Public Lc3 As ADODB.Recordset         'recordset com as configurações padrão, como som de alarme etc..
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'variaiveis globais
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'variaveis de acesso
Public vidusuariolib As Long          'id do usuario liberado para acesso
Public vnickusuario As String         'nick do usuario liberado para acesso
Public vsncadastrada As String        'senha cadastrada para login
Public vsndigitada As String          'senha digitada para login

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'abertura de conexao
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'obtem o caminho a base de dados padrão
Public Sub ObtemCaminhoBase()
    'obtem o tipo de base (alterar para o local do windows)
    vTip = ReadINI("Base", "Tipo", "C:\lan\Arquivos\Paramt.ini")
    'verifica o caminho a ser carregado
    If vTip = "D" Then
        'ACCESS - carrega o caminho ao banco de dados obtido no arquivo ini
        sCamin1 = ReadINI("CaminhoBaseDTA", "caminho", "C:\lan\Arquivos\Paramt.ini")
    ElseIf vTip = "F" Then
        'FIREBIRD - carrega o caminho ao banco de dados obtido no arquivo ini
        sCamin1 = ReadINI("CaminhoBaseFIR", "caminho", "C:\lan\Arquivos\Paramt.ini")
    End If
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'ADO
Public Sub AbreCon()
    'chama o cmainho
    Call ObtemCaminhoBase
    'verifica o tipo de conexao Access/Dta ou Firebird
    If vTip = "F" Then
        'conexao Firebird (ODBC)
        con.Open ("DRIVER=Firebird/InterBase(r) driver; UID=SYSDBA;PWD=BUCETA; DBNAME=" & sCamin1)
    ElseIf vTip = "D" Then
        'conexao Access
        con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sCamin1 & ";Mode=ReadWrite;Persist Security Info=False;Jet OLEDB:Database Password=matweb146987syswm770"
    End If
End Sub
Public Sub FechaCon()
    con.Close
    Set con = Nothing
End Sub
