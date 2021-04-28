VERSION 5.00
Begin VB.Form frmOpcadastro 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Super Imob - Opções"
   ClientHeight    =   1485
   ClientLeft      =   4605
   ClientTop       =   3105
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3240
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imóveis"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pessoas"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmOpcadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Option1.Value = False
Option2.Value = False
End Sub
