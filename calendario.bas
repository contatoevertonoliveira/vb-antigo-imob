Attribute VB_Name = "calendario"
'-- CONSTANTES
Public Const modal = 1

'-- CORES
Public Const Amarelo = &HFFFF&
Public Const Branco = &HFFFFFF
Public Const Titulo = "Dia-A-Dia"

'-- VARIÁVEIS
Public IndMes As Integer    'Índice mês: 1 a 12
Public NomeMes As String    'Nome do mês
Public NumDias As Integer   'Número do dia do mês
Public ANO As Integer       'Ano-calendário
Public AnoMom As Integer    'Ano atual no programa
Public MesMom As Integer    'Mês idem
Public DiaDaSemana As Integer '1 a 7 (dom a sab)

'-- Constantes - teclas
Public Const KEY_LEFT = &H25    'Seta esquerda
Public Const KEY_UP = &H26        'Seta para cima
Public Const KEY_RIGHT = &H27    'Seta direita
Public Const KEY_DOWN = &H28    'Seta para baixo

Sub CalculaDias()
'-- Calcula a posição do dia 1o. do mês
'-- e determina o dia da semana correspondente
MM% = IndMes
DD% = 1
YY% = ANO
    Relval# = DateSerial(YY%, MM%, DD%)
    DiaDaSemana = Weekday(Relval#)
    EscreveMes
End Sub

Sub EscreveMes()
'-- Escreve p titulo: mês e ano
frmCalendario.TitMesAno.Caption = NomeMes + " " + Str$(ANO)

'-- Limpa o calendário para evitar sobreposição de
'-- informações de um mês para outro. Pinta todas a
'-- casas de branco; isso evita que casa não ocupada
'-- (virada de mês) fique pintada de amarelo
For N% = 0 To 41
    frmCalendario.Casa(N%).Caption = " "
    frmCalendario.Casa(N%).BackColor = Branco
Next N%

'-- Escreve as datas nas células
p% = DiaDaSemana
pp% = p% - 2   'dif. entre data e sua posição
                'em Casa(index) index=0 a 41
                
'-- Esclarecimento de pp*:
'-- n* (o num de dias) vai de 1 a 31; p* vai de 1 a 7
'-- posição dia 1 na Casa(index): Casa(p*-1)
'-- posição dia n* na Casa(index): Casa(n*+(p*-1)-1),
'-- ou Casa(n*+p*-2). Portanto, pp*=p*-2

For N% = 1 To NumDias
    ' Se mês atual, mostra o dia atual em amarelo
    frmCalendario.Casa(p% - 1).Caption = Str$(N%)
    If IndMes = Month(Now) And ANO = Year(Now) And N% = Day(Now) Then
        frmCalendario.Casa(N% + pp%).BackColor = Amarelo
    Else
        frmCalendario.Casa(N% + pp%).BackColor = Branco
    End If
    p% = p% + 1
Next N%

Unload frmSelMes       'Descarrega frmselmes
MesMom = IndMes         'Guarda o mês vigente
    
End Sub

Sub CalculaMes()
'-- Informa ao programa o nome dos meses em português.

    Select Case IndMes
        Case 1
            NomeMes = "Janeiro"
            NumDias = 31
        Case 2
            NomMes = "Fevereiro"
            'Bissexto é ano múltiplo de 4 e não de 100; ou múltiplo de 400
            If (ANO Mod 4 = 0 And ANO Mod 100 <> 0) Or (ANO Mod 400 = 0) Then
                NumDias = 29
            Else
                NumDias = 28
            End If
        Case 3
            NomeMes = "Março"
            NumDias = 31
        Case 4
            NomeMes = "Abril"
            NumDias = 30
        Case 5
            NomeMes = "Maio"
            NumDias = 31
        Case 6
            NomeMes = "Junho"
            NumDias = 30
        Case 7
            NomeMes = "Julho"
            NumDias = 31
        Case 8
            NomeMes = "Agosto"
            NumDias = 31
        Case 9
            NomeMes = "Setembro"
            NumDias = 30
        Case 10
            NomeMes = "Outubro"
            NumDias = 31
        Case 11
            NomeMes = "Novembro"
            NumDias = 30
        Case 12
            NomeMes = "Dezembro"
            NumDias = 31
    End Select
    CalculaDias
               
End Sub

Sub ErroAno()

'-- Mensagem de erro para ano inválido
M$ = "Ano inválido. O ano deve estar no " + Chr$(13) + Chr$(10)
M$ = M$ + "intervalo de 1753 a 2078, inclusive."
MsgBox M$, 48, Titulo

End Sub


Sub LeDiadaSemana()
'-- ensina ao programa os nomes dos dias da
'-- semana em português a partir de n* (1 a 7),
'-- índice do dia lido no relógio da máquina.

N% = Weekday(Now)
Select Case N%
    Case 1
        a$ = "domingo"
    Case 2
        a$ = "segunda-feira"
    Case 3
        a$ = "terça-feira"
    Case 4
        a$ = "quarta-feira"
    Case 5
        a$ = "quinta-feira"
    Case 6
        a$ = "sexta-feira"
    Case Else
        a$ = "sábado"
End Select
frmCalendario.JanData.Caption = Format$(Now, "dd/mm/yyyy") + ", " + a$

End Sub

Sub MesAnt()
'-- Define saltos de jan//dez, dez/jan
    If IndMes = 1 Then
        If ANO = 1753 Then
            ErroAno
            Exit Sub
        Else
            ANO = ANO - 1
            IndMes = 12
        End If
    Else
        IndMes = IndMes - 1
    End If
    CalculaMes
    
End Sub

Sub MesAtual()
    IndMes = Month(Now)     'Captura mês atual(1 a 12)
    ANO = Year(Now)         'Captura ano atual
    LeDiadaSemana           'Atualiza JanData
    CalculaMes              'Entra no fluxo normal do programa
End Sub


Sub MesProx()
'-- Define saltos de jan/dez, dez/jan
    If IndMes = 12 Then
        If ANO = 2078 Then
            ErroAno
            Exit Sub
        Else
            ANO = ANO + 1
            IndMes = 1
        End If
    Else
        IndMes = IndMes + 1
    End If
    CalculaMes
    
End Sub

Sub SelectText(TBox As Control)
'-- Em caso de erro, faz voltar à janela,
'-- selecionando o texto
TBox.SetFocus
TBox.SelStart = 0
TBox.SelLength = Len(TBox.Text)
End Sub


