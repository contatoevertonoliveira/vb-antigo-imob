VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Super Imob - Menu"
   ClientHeight    =   4905
   ClientLeft      =   2580
   ClientTop       =   1935
   ClientWidth     =   7470
   ForeColor       =   &H00000000&
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7470
   Begin VB.CommandButton Command1 
      Caption         =   "Finalizar Menu / Sair"
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   4200
      Width           =   3015
   End
   Begin VB.CommandButton cmdCalen 
      Caption         =   "&Calendário"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdImpressora 
      Caption         =   "&Impressora"
      Enabled         =   0   'False
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdRec 
      Caption         =   "- &Emissão de Recibos -"
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton cmdConsulta 
      Caption         =   "- &Consultas e Pesquisas -"
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "- &Calculadora -"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdAlerta 
      Caption         =   "- Em &Alerta -"
      Enabled         =   0   'False
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdAgenda 
      Caption         =   "- &Agenda de Compromissos -"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton cmdCadastros 
      Caption         =   "- &Cadastros -"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdContratos 
      Caption         =   " - &Contratos -"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   6975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAINEL CENTRAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   480
      Left            =   1920
      TabIndex        =   10
      Top             =   240
      Width           =   3600
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgenda_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAgenda.Font.Bold = True
End Sub

Private Sub cmdAlerta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAlerta.Font.Bold = True
End Sub

Private Sub cmdCadastros_Click()
frmEscolha.Show
End Sub

Private Sub cmdCadastros_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCadastros.Font.Bold = True
End Sub

Private Sub cmdCalc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCalc.Font.Bold = True
End Sub

Private Sub cmdCalen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCalen.Font.Bold = True
End Sub

Private Sub cmdConsulta_Click()
frmBusca.Show
Unload frmMenu
End Sub

Private Sub cmdConsulta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdConsulta.Font.Bold = True
End Sub

Private Sub cmdContratos_Click()
frmOpContrato.Show
End Sub

Private Sub cmdContratos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdContratos.Font.Bold = True
End Sub

Private Sub cmdImpressora_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdImpressora.Font.Bold = True
End Sub

Private Sub cmdRec_Click()
frmConRec.Show
End Sub

Private Sub cmdRec_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdRec.Font.Bold = True
End Sub

Private Sub Command1_Click()
If MsgBox("Quer sair do Menu Principal?", vbYesNo, "Sair do Menu") = vbYes Then
    Unload Me
Else
    Exit Sub
End If
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Font.Bold = True
End Sub

Private Sub Form_Activate()
Unload frmShow
MousePointer = 0 ' cursor normal
Me.BackColor = &H8000000F
Frame1.BackColor = &H8000000F
End Sub

Private Sub Form_Deactivate()
Me.BackColor = &H80000009
Frame1.BackColor = &H80000009
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Height) / 2
Me.Left = (Screen.Width - Width) / 2
MousePointer = 12 'Ampulheta
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Font.Bold = False
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdContratos.Font.Bold = False
cmdCadastros.Font.Bold = False
cmdAgenda.Font.Bold = False
cmdAlerta.Font.Bold = False
cmdCalc.Font.Bold = False
cmdCalen.Font.Bold = False
cmdConsulta.Font.Bold = False
cmdImpressora.Font.Bold = False
cmdRec.Font.Bold = False
Command1.Font.Bold = False

End Sub
