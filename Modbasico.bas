Attribute VB_Name = "Modbasico"
'leitura do arquivo INI
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'gravação do arquivo INI
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'uso do Internet Explorer e outlok express uso do form EMP1
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'declaração para desabilitar o ctrl + alt + del
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'declaração das variaveis para obter o serial do hd
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public vhds1 As Long                         'numero de serie do H.D.
Public vhdn1 As String                       'nome do H.D.
Public vhdt1 As String                       'tipo de H.D.
Public vcpn1 As String * 255                 'nome do computador
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'contadores de acesso e serial
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public vcontador As String      'contador atual
Public vnovovalor As String     'novo valor para o contador
Public valorusuario As String   'valor a ser informado ao usuario
Public vDtatual As Date         'data atual do sistema
Public vDtgravada As Date       'data gravada
Public vAcescontador As String  'contador de acesso no dia
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'gerador da senha de acesso (verificador)
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public vregistro1 As String    'abs numero do hd
Public vregistro2 As String    'mes da data
Public vregistro3 As String    'ano da data
Public vregistro4 As String    'realiza a montagem
Public vregistro5 As Long      'realiza a primeira soma
Public vregistro6 As Long      'realiza a primeira soma
Public vregistro7 As String    'realiza o procedimento hexa
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'contador de tempo de acesso e uso do sistema
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public iCotMin As Long             'contador de segundos em uso
Public iInic1 As Long              'contador inicial carregado - segundos da hora
Public iInic2 As Long              'contador inicial carregado - minutos transformados em segundos
Public iInic3 As Long              'contador inicial carregado - horas transformados em segundos
Public iLimt1 As Long              'somatorio dos segundos carregado no acesso do sistema
Public iLimt2 As Long              'limite de tempo de uso (4,5 horas) marca um contador
Public iLimt3 As Long              'limite de tempo de uso (9 horas) marca um contador
Public iLimt4 As Long              'limite de tempo de uso (14 horas) avisa o usuario
Public iLimt5 As Long              'limite final (10 minutos depois do iLimt4 - o sistema é encerrado
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'arquivos ini
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'leitura do INI
Public Function ReadINI(Secao As String, Entrada As String, Arquivo As String)
Dim retlen As String
Dim Ret As String
    Ret = String$(255, 0)
    retlen = GetPrivateProfileString(Secao, Entrada, "", Ret, Len(Ret), Arquivo)
    Ret = Left$(Ret, retlen)
    ReadINI = Ret
End Function
'gravação do INI
Public Sub WriteINI(Secao As String, Entrada As String, Texto As String, Arquivo As String)
    WritePrivateProfileString Secao, Entrada, Texto, Arquivo
End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'criptografia
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Function Crypt(Text As String) As String
Dim strTempChar As String  'montagem da criptografia
Dim i As Integer
    'montagem a criptografia
    For i = 1 To Len(Text)
        If Asc(Mid$(Text, i, 1)) < 128 Then
            strTempChar = Asc(Mid$(Text, i, 1)) + 128
        ElseIf Asc(Mid$(Text, i, 1)) > 128 Then
            strTempChar = Asc(Mid$(Text, i, 1)) - 128
        End If
        Mid$(Text, i, 1) = Chr(strTempChar)
    Next i
    Crypt = Text
End Function
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'geração de backup automatico
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'função de geração de backup
Public Sub GeraBackupsDados()
On Error GoTo errobackup
Dim vCOMPCT As String   'caminho e banco de dados compactado
Dim vDIA As String      'dia da data de compactaçao
Dim vMES As String      'mes da data de compactação
Dim vANO As String      'ano da data de compactação

    'obtendo a montagem da data
    vDIA = Day(Date)
    vMES = Month(Date)
    vANO = Year(Date)
    
   
    'monta o novo caminho e o nome do bakup
    vCOMPCT = "C:" + "\backup" + "-" + vDIA + "-" + vMES + "-" + vANO & ".bkp"
        
    'chama o cmainho
    Call ObtemCaminhoBase
    
    'verifica o tipo de conexao Access/Dta ou Firebird
    '************************
    'se for um banco FIREBIRD
    '************************
    If vTip = "F" Then
        DoEvents
        'copia o banco para o drive selecionado
        FileCopy sCamin1, vCOMPCT
    '*********************
    'se for um banco ACESS
    '*********************
    ElseIf vTip = "D" Then
        DoEvents
        'compacta o arquivo para gerar o backup
        DBEngine.CompactDatabase sCamin1, vCOMPCT, dbLangGeneral & ";pwd=matweb146987syswm770", dbEncrypt, ";pwd=matweb146987syswm770"
        'compacta o arquivo para gerar a nova base compactada
        DBEngine.CompactDatabase sCamin1, "C:\basecompactada.dtl", dbLangGeneral & ";pwd=matweb146987syswm770", dbEncrypt, ";pwd=matweb146987syswm770"
        'elimina a base antiga
        Kill sCamin1
        'renomeia a nova base compactada
        Name "C:\basecompactada.dtl" As sCamin1
    End If
   
    'mensagem ao usuario
    MsgBox "Realizado o backup com sucesso ...", , "Backup de dados"
      
'erros variados pela escolha do drive
errobackup:
If Err.Number = 52 Then
    MsgBox "A Base de dados não está localizada nessa máquina ...", , "Backup cancelado"
    Resume fimbackup
ElseIf Err.Number = 53 Then
    MsgBox "A Base de dados não está localizada nessa máquina ...", , "Backup cancelado"
    Resume fimbackup
ElseIf Err.Number = 61 Then
    MsgBox "Espaço indisponível no drive selecionado ...", , "Backup cancelado"
    Resume fimbackup
ElseIf Err.Number = 70 Then
    MsgBox "Existem usuários utilizando o banco de dados ...", , "Backup cancelado"
    Resume fimbackup
ElseIf Err.Number = 71 Then
    MsgBox "Insira um disquete no drive selecionado ...", , "Backup cancelado"
    Resume fimbackup
ElseIf Err.Number = 76 Then
    MsgBox "Não foi possível realizar o backup no drive selecionado ...", , "Backup cancelado"
    Resume fimbackup
ElseIf Err.Number = 3204 Then
    MsgBox "O backup diário já foi realizado, exclua backup" & vbCrLf & " existente e efetue a operação novamente ...", , "Backup cancelado"
    Resume fimbackup
ElseIf Err.Number = 3356 Then
    MsgBox "Existem usuários utilizando o banco de dados ...", , "Backup cancelado"
    Resume fimbackup
End If

fimbackup:
    
End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'conferencia cnpj/cpf
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Function FormataCIC(Q_Cic As String) As String
  FormataCIC = IIf(Len(Q_Cic) > 11, Format(Q_Cic, "@@.@@@.@@@/@@@@-@@"), Format(Q_Cic, "@@@.@@@.@@@-@@"))
End Function
Public Function LimpaCampo(Q_Campo) As String
  Dim CampoLimpo As String
  Dim X As Integer
  
  For X = 1 To Len(Q_Campo)
    If Mid(Q_Campo, X, 1) <> "/" And Mid(Q_Campo, X, 1) <> "-" And Mid(Q_Campo, X, 1) <> "." And Mid(Q_Campo, X, 1) <> ":" And Mid(Q_Campo, X, 1) <> "," And Mid(Q_Campo, X, 1) <> "_" And Mid(Q_Campo, X, 1) <> "," Then CampoLimpo = CampoLimpo & Mid(Q_Campo, X, 1)
  Next X
  
  LimpaCampo = IIf(Len(CampoLimpo) = 0, "", CampoLimpo)
End Function
Public Function Texto_MainFrame(QTexto As String) As String
  Dim Y As Integer

  For Y = 1 To Len(QTexto)
    Select Case Mid(QTexto, Y, 1)
    Case "á", "ã", "â", "Á", "Ã", "Â"
      Mid(QTexto, Y, 1) = "A"
    Case "ç", "Ç"
      Mid(QTexto, Y, 1) = "C"
    Case "é", "ê", "É", "Ê"
      Mid(QTexto, Y, 1) = "E"
    Case "í", "î", "Í", "Î"
      Mid(QTexto, Y, 1) = "I"
    Case "ó", "õ", "ô", "Ó", "Õ", "Ô"
      Mid(QTexto, Y, 1) = "O"
    Case "ú", "û", "Ú", "Û"
      Mid(QTexto, Y, 1) = "U"
    End Select
  Next Y
  
  Texto_MainFrame = UCase(QTexto)
End Function
Public Function Valida_CIC(Q_Cic As String) As Boolean
  Q_Cic = LimpaCampo(Q_Cic)
  
  Select Case Len(Q_Cic)
  Case Is < 11
    Valida_CIC = False
  Case Is > 14
    Valida_CIC = False
  Case 12 To 14
    Valida_CIC = ValidaCGC(Q_Cic)
  Case Else
    Valida_CIC = ValidaCPF(Q_Cic)
  End Select
End Function
Public Function ValidaCPF(CPF As String) As Boolean
  Dim Digito, X, Soma, Resto As Integer, Regua As Variant
     
  If Len(Trim(CPF)) < 11 Then
    ValidaCPF = False
    Exit Function
  End If
  'Calcula o primeiro dígito verificador
  Regua = Array(11, 10, 9, 8, 7, 6, 5, 4, 3, 2)
  For X = 1 To 9
    Soma = Soma + (Regua(X) * Mid(CPF, X, 1))
  Next X
  Resto = Soma Mod 11
  Digito = IIf(Resto = 0 Or Resto = 1, "0", Trim(11 - Resto))
  'Calcula o segundo dígito verificador
  Soma = 0
  Resto = 0
  For X = 0 To 9
    Soma = Soma + (Regua(X) * Mid(CPF, X + 1, 1))
  Next X
  Resto = Soma Mod 11
  Digito = Digito + IIf(Resto = 0 Or Resto = 1, "0", Trim(11 - Resto))
  ValidaCPF = IIf(Mid(CPF, 10, 2) <> Digito, False, True)
End Function
Public Function ValidaCGC(CGC As String) As Boolean
  'Valida primeiro digito
  ValidaCGC = TestaDig(Left(CGC, 13), 11)
  'Valida segundo digito
  If ValidaCGC = True Then
    ValidaCGC = TestaDig(Trim(CGC), 11)
  End If
End Function
Public Function TestaDig(Q_Dado As String, Q_Base As Integer) As Boolean
  If Q_Base = 10 Then
    'Testa dígito na base 10
    TestaDig = IIf(Base_10(Left(Trim(Q_Dado), Len(Trim(Q_Dado)) - 1)) <> Right(Trim(Q_Dado), 1), False, True)
  ElseIf Q_Base = 11 Then
    'Testa dígito na base 11
    TestaDig = IIf(Base_11(Left(Trim(Q_Dado), Len(Trim(Q_Dado)) - 1)) <> Right(Trim(Q_Dado), 1), False, True)
  End If
End Function
Public Function Base_10(Q_Dado) As String
  Dim DadoCalc, Peso, Soma, Resto, X As Integer, Regua As String
  DadoCalc = LimpaCampo(Q_Dado)
  Peso = 2
  For X = Len(DadoCalc) To 1 Step -1
    Regua = Regua + Trim((Mid(DadoCalc, X, 1) * Peso))
    Peso = IIf(Peso = 1, 2, 1)
  Next X
  X = 1
  For X = 1 To Len(Regua)
    Soma = Soma + Val(Mid(Regua, X, 1))
  Next X
  Resto = Soma Mod 10
  Base_10 = Right(Trim(10 - Resto), 1)
End Function
Public Function Base_11(QNumero As String) As String
  Dim Numero, i, Produto, Multiplicador, Digito As Integer

  Numero = Trim(QNumero)
  'Calcula digito do modulo 11
  Multiplicador = 2
  For i = Len(Numero) To 1 Step -1
    Produto = Produto + Val(Mid(Numero, i, 1)) * Multiplicador
    Multiplicador = IIf(Multiplicador = 9, 2, Multiplicador + 1)
  Next i
  'Exceção
  Digito = 11 - Int(Produto Mod 11)
  Digito = IIf(Digito = 10 Or Digito = 11, 0, Digito)
  Base_11 = Trim(Str(Digito))
End Function
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'função de contador de acesso de uso ao sistema
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Sub ContadorTempoUso()
    DoEvents
    'define o contador
    iCotMin = iCotMin + 1
    'limite 4,5 horas
    If iCotMin = iLimt2 Then
        'grava o valor de acesso
        Call GravacaoContadorDiario
    'limite 9 horas
    ElseIf iCotMin = iLimt3 Then
        'grava o valor de acesso
        Call GravacaoContadorDiario
    'limite 14 horas
    ElseIf iCotMin = iLimt4 Then
        'grava o valor de acesso
        Call GravacaoContadorDiario
        'avisa o usuario q dentro de 10 minutos o sistema encerrara
        MsgBox "Para proteção dos dados, finalize as tarefas pendentes," & vbCrLf & "encerre o aplicativo e se desejar acesse-o novamente ...", vbCritical, "O sistema será encerrado em 10 minutos"
    'limite 14 horas
    ElseIf iCotMin = iLimt5 Then
        'fecha a conexao com o banco
        Call FechaCon
        'encerra o sistema
        End
    End If
End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'função que analisa o contador para realizar a gravação do novo valor
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Sub GravacaoContadorDiario()
Dim vNewAcescontador As String
    vNewAcescontador = ReadINI("config4", "Contador", "C:\Windows\SysConfig.tts")
    'verifica o numero de contagens
    Select Case vNewAcescontador
    Case "A01"
        'grava o novo contador de sequencia diario
        Call WriteINI("config4", "Contador", "B02", "C:\Windows\SysConfig.tts")
    Case "B02"
        'grava o novo contador de sequencia diario
        Call WriteINI("config4", "Contador", "C03", "C:\Windows\SysConfig.tts")
    Case "C03"
        'grava o novo contador de sequencia diario
        Call WriteINI("config4", "Contador", "D04", "C:\Windows\SysConfig.tts")
    Case "D04"
        'grava o novo valor criptografado no arquivo ini
        Call WriteINI("config1", "control", Crypt(vnovovalor), "C:\Windows\SysConfig.tts")
        'grava a data de ultimo acesso
        Call WriteINI("config3", "DataDia", Crypt(Format(Date, "Short Date")), "C:\Windows\SysConfig.tts")
        'zera o contador de sequencia diario
        Call WriteINI("config4", "Contador", "A01", "C:\Windows\SysConfig.tts")
    End Select
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'abertura de conexao
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'obtem o caminho a base de dados padrão
Public Sub DefiniTipoTravamentoTeclado()
    'obtem o tipo de base (alterar para o local do windows - USAR XP OU 98)
    vTrvTec = ReadINI("Teclado", "TravaTeclado", "C:\lan\Arquivos\Paramt.ini")
    'se for 98
    If vTrvTec = "98" Then
        CtrlAltDel True
    'se for XP
    ElseIf vTrvTec = "XP" Then
        Open "c:\windows\system32\taskmgr.exe" For Binary As #1
    End If
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'procedimento para desabilitar o ctrl + alt + del (PARA WIN98)
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub CtrlAltDel(Desabilita As Boolean)
  Call SystemParametersInfo(97, Desabilita, "1", 0)
End Sub
