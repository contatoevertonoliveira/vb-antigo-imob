VERSION 5.00
Begin VB.Form frmSelMes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mês e Ano"
   ClientHeight    =   2475
   ClientLeft      =   3135
   ClientTop       =   2760
   ClientWidth     =   4800
   Icon            =   "SelecionaMes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2475
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox JanAno 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   315
      Left            =   3300
      MaxLength       =   4
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton BotaoCanc 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3360
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton BotaoOk 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3360
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Digite o ano:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3300
      TabIndex        =   16
      Top             =   150
      Width           =   1515
   End
   Begin VB.Label Label13 
      Caption         =   "Escolha o mês:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   225
      TabIndex        =   15
      Top             =   150
      Width           =   1890
   End
   Begin VB.Label LabMes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dezembro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   11
      Left            =   1575
      TabIndex        =   14
      Top             =   2025
      Width           =   1215
   End
   Begin VB.Label LabMes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Novembro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   10
      Left            =   1575
      TabIndex        =   13
      Top             =   1725
      Width           =   1215
   End
   Begin VB.Label LabMes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Outubro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   9
      Left            =   1575
      TabIndex        =   12
      Top             =   1425
      Width           =   1215
   End
   Begin VB.Label LabMes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Setembro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   8
      Left            =   1575
      TabIndex        =   11
      Top             =   1125
      Width           =   1215
   End
   Begin VB.Label LabMes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agosto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   7
      Left            =   1575
      TabIndex        =   10
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label LabMes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Julho"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   6
      Left            =   1575
      TabIndex        =   9
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label LabMes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Junho"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   5
      Left            =   225
      TabIndex        =   8
      Top             =   2025
      Width           =   1215
   End
   Begin VB.Label LabMes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Maio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   4
      Left            =   225
      TabIndex        =   7
      Top             =   1725
      Width           =   1215
   End
   Begin VB.Label LabMes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Abril"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   3
      Left            =   225
      TabIndex        =   6
      Top             =   1425
      Width           =   1215
   End
   Begin VB.Label LabMes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Março"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   225
      TabIndex        =   5
      Top             =   1125
      Width           =   1215
   End
   Begin VB.Label LabMes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fevereiro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   225
      TabIndex        =   4
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label LabMes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Janeiro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   225
      TabIndex        =   3
      Top             =   525
      Width           =   1215
   End
End
Attribute VB_Name = "frmSelMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----------------------------------------------
' frmselmes.FRM
' Permite a escolha de ano e mês quaisquer
' ----------------------------------------------

' Variáveis válidas só para o form. Inserir
' na área de General Declarations do form
Dim Reserva As Integer
Dim Setas As Integer

Sub Teste()

'-- Captura ano digitado pelo usuário;
'-- testa esse ano; se inválido, chama
'-- mensagem de erro; recupera o ano
'-- que era vigente (AnoMom) antes do erro

ANO = Val(JanAno.Text)      'lê o ano

'-- No VB 2.0, ano máximo é 9999
If ANO >= 1753 And ANO <= 2078 Then     'teste
    CalculaMes
Else
    ErroAno
    SelectText JanAno       'Põe o foco na janela Ano
    ANO = AnoMom            'Recupera ano vigente
End If

End Sub


Private Sub BotaoCanc_Click()

'-- Operação cancelada; portanto, o mês ainda
'-- é o mesmo que está mostrado no calendário.
'-- Se frmselmes for chamado logo em seguida,
'-- mostrará esse mês em amarelo

IndMes = MesMom     'Recupera o mês
Unload frmSelMes

End Sub

Private Sub BotaoOk_Click()
    AnoMom = ANO    'Guarda ano vigente
    Teste
End Sub

Private Sub Form_Load()
CENTRALIZA_FORM Me
'Cursor volta ao normal
Screen.MousePointer = 0
'Pinta de amarelo a label com o nome do mês
LabMes(IndMes - 1).BackColor = Amarelo
End Sub


Private Sub JanAno_KeyDown(KeyCode As Integer, Shift As Integer)
'-- Muda o mês com teclas para cima ou para baixo
'-- Variável Reserva guarda o mês vigente antes da tecla
'-- Os Ifs controlam saltos dez/jan e jan/dez

Reserva = IndMes

If KeyCode = KEY_UP Then    'tecla p/ cima
    If IndMes > 1 Then
        IndMes = IndMes - 1
    Else
        IndMes = 12
    End If
ElseIf KeyCode = KEY_DOWN Then  'tecla p/ baixo
    If IndMes < 12 Then
        IndMes = IndMes + 1
    Else
        IndMes = 1
    End If
Else
End If

'-- Setas indica que uma tecla foi pressionada
'-- Usa rotina do clique do mouse sobre o mês
Setas = True
LabMes_Click IndMes - 1

End Sub


Private Sub LabMes_Click(Index As Integer)

'-- IndMes (em j%=...) indica o mês vigente
'-- antes desta rotina. Setas diz que uma
'-- tecla de direção foi acionada
'-- Reserva é o IndMes vigente antes
'-- do acionamento da tecla

    j% = IndMes - 1
    
    If Setas Then
        Setas = False
        j% = Reserva - 1
    End If
    
    LabMes(j%).BackColor = Branco
    LabMes(Index).BackColor = Amarelo
    
    IndMes = Index + 1
    
End Sub


Public Sub CENTRALIZA_FORM(Formulario As Form)
On Error Resume Next 'Evita erro caso o usuário minimize o Form
With Formulario
    .Left = (Screen.Width - .Width) / 2 'Alinha o form no horizontalmente no centro
    .Top = (Screen.Height - .Height) / 2  'Alinha o form no verticalmente no centro
End With
'With Formulario
'    .Left = ((mdiGerest.Width - .Width) / 2) 'Alinha o form no horizontalmente no centro
'    .Top = ((mdiGerest.Height - .Height) / 2) - 1000 'Alinha o form no verticalmente no centro
'End With
End Sub


