VERSION 5.00
Begin VB.Form Lanaviso 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer DuracaoMessagem 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   600
      Top             =   2760
   End
   Begin VB.PictureBox FundoMessagem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C000&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1665
      ScaleWidth      =   5925
      TabIndex        =   0
      Top             =   0
      Width           =   5955
      Begin VB.Label Mensagem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "XX Mensagem XX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   5655
         WordWrap        =   -1  'True
      End
      Begin VB.Label cronometro 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2280
         TabIndex        =   1
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   5640
         MouseIcon       =   "Lanaviso.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Lanaviso.frx":0CCA
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Timer ControleFechamento 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1080
      Top             =   2760
   End
   Begin VB.Timer ControleAbertura 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   2760
   End
   Begin VB.Timer TimerContador 
      Interval        =   1
      Left            =   1560
      Top             =   2760
   End
End
Attribute VB_Name = "Lanaviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'declaração das apis
Private Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex As Long)
Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'declarações para manter o form sempre no topo
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'constantes da janela
Const SM_CXFULLSCREEN = 16   'posição inicial da janela em X no monitor
Const SM_CYFULLSCREEN = 17   'posição inicial da janela em Y no monitor
Const SND_SYNC = &H0         'variaveis para o uso de som
Const SND_ASYNC = &H1        'possibilitando o som ser carregado conforme
Const SND_NODEFAULT = &H2    'a tela surge no evento DysplayAlert
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10
'constantes para manter o form sempre no topo
Const HWND_TOPMOST = -1
Const SWP_SHOWWINDOW = &H40
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'procedimento de definição e abertura (load/show) do form messagem
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Sub DisplayAlert(Messagem As String, Duracao As Long)
Dim wFlags As Long, X As Long
    'contadores da janela
    AlertCount = AlertCount + 1
    AlertIndex = AlertCount
    'carrega a variavel com a mensagem a ser exibida
    Mensagem.Caption = Messagem
    'define o tempo que será utilizado no controle timer
    'de tempo de visualização da mensagem
    DuracaoMessagem.Interval = Duracao
    'obtem as medidas utilizadas na tela
    Tx = GetSystemMetrics(SM_CXFULLSCREEN)            'medida em pixels em X(horizontal)
    Ty = GetSystemMetrics(SM_CYFULLSCREEN)             'medida em pixels em Y(vertical)
    lngScaleX = Me.Width - Me.ScaleWidth               'definição das escalas
    lngScaleY = Me.Height - Me.ScaleHeight
    'definição do tamanho e posição do form em relacção ao picture de fundo
    Me.Height = 90                                     'altura
    Me.Width = FundoMessagem.Width + lngScaleX  'largura
    Me.Left = Tx * Screen.TwipsPerPixelX - Me.Width    'posição
    Me.Top = (Ty * Screen.TwipsPerPixelY) - ((FundoMessagem.Height + lngScaleY) * (AlertCount - 1)) + 300
    'carrega o form de messagem
    Me.Show
    'libera o timer de controle de abertura/carregamento
    ControleAbertura.Enabled = True
    
    'se quiser usar sinal sonoro
    wFlags = SND_ASYNC Or SND_NODEFAULT
    X = sndPlaySound(vcaminhodosom, wFlags)
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'mantem o form no topo (popup)
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Sub Form_Load()
Dim retValue As Long
    retValue = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 22000, 22000, SWP_SHOWWINDOW)
End Sub
'TIMERS
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'procedimento de carregamento do form com limite ate o picturebox de fundo
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Sub ControleAbertura_Timer()
Dim PosFinal As Long       'posição final(total) do form
Dim PosInicial As Long     'posição inicial do form
    'obtem a posição final a ser carregada
    PosFinal = Me.Height
    'enquanto nao chegar ao limite(do picute) vai carregando
    If PosFinal < FundoMessagem.Height + lngScaleY Then
        'vai carregando a tela (subindo)
        PosFinal = PosInicial + 30
        Me.Height = Me.Height + (PosFinal - PosInicial)
        Me.Top = Me.Top - (PosFinal - PosInicial)
    Else
        'ao completar a subida (a tela estiver toda visualizada)
        'bloqueia o procedimento de abertura/carregamento
        ControleAbertura.Enabled = False
        'libera o procedimento de controle de tempo de visualização
        DuracaoMessagem.Enabled = True
        'torna visivel o cronometro(label com o relogio)
        cronometro.Visible = True
    End If
End Sub
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'procedimento que controla o tempo em que o form ficará sendo visualizado
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Sub DuracaoMessagem_Timer()
    'bloqueia o timer de controle do tempo da messagem
    DuracaoMessagem.Enabled = False
    'aciona o timer que controla o evento de fechamento da mensagem
    ControleFechamento.Enabled = True
End Sub
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'procedimento de fechamento do form
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Sub ControleFechamento_Timer()
    'fecha o form
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'procedimento para carregar o relogio da messagem
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Sub TimerContador_Timer()
    'carrega o cronometro
    cronometro.Caption = Time
End Sub
'fecha
Private Sub Image1_Click()
    Me.Hide
End Sub
