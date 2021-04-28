Attribute VB_Name = "variaveis"
Public vcaminhodosom As String        'caminho ao arquivo do som para o form de alerta/aviso
'chama a função
Public Sub AcionaAviso(Text As String)
Dim AlertBox As Lanaviso
    'cria uma nova instancia do objeto
    Set AlertBox = New Lanaviso
    'chama o procedimento do form, carregando as variaveies da mensagem e da duração do evento
    AlertBox.DisplayAlert Text, 16000
End Sub
