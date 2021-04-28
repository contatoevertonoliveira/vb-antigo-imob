VERSION 5.00
Begin VB.Form frmOpContrato 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Super Imob - Opções"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmOpContrato.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Escolha o tipo de Busca:"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      Begin VB.OptionButton Option5 
         Caption         =   "Pessoas"
         Height          =   195
         Left            =   3000
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Contrato de Compra e Venda"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   2415
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Contrato de Locação"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Clientes"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmOpContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
If Option1.Value = True Then
Unload frmOpContrato
frmPesq.Show 1
End If
If Option2.Value = True Then
Unload frmOpContrato
frmPesContr.Show 1
End If

If Option1.Value = False Then
If Option2.Value = False Then
If Option4.Value = False Then
If Option5.Value = False Then
Unload Me
End If
End If
End If
End If
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Height) / 2
Me.Left = (Screen.Width - Width) / 2
Option1.Value = False
Option2.Value = False
Option4.Value = False
Option5.Value = False
End Sub
