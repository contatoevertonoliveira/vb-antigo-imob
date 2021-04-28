VERSION 5.00
Begin VB.Form frmEscolha 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Escolha"
   ClientHeight    =   2685
   ClientLeft      =   4605
   ClientTop       =   2175
   ClientWidth     =   2640
   Icon            =   "frmEscolha.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   2640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   735
      Left            =   720
      Picture         =   "frmEscolha.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Escolha uma das opções:"
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2175
      Begin VB.OptionButton Option2 
         Caption         =   "Clientes"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Proprietários"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmEscolha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Option1.Value Then
    FrmProprietarios.Show
    Unload Me
End If
If Option2.Value Then
    frmClientes.Show
    Unload Me
End If
If Option1.Value = False Then
If Option2.Value = False Then
Unload frmEscolha
End If
End If
End Sub

Private Sub Form_Load()
Option1.Value = False
Option2.Value = False
End Sub

Private Sub Option1_Click()
If Option1.Value Then
    Command1.Caption = "Propriet&ários"
End If
End Sub

Private Sub Option2_Click()
If Option2.Value Then
    Command1.Caption = "Cli&entes"
End If
End Sub
