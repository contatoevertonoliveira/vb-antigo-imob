VERSION 5.00
Begin VB.Form Contador 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contador"
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2070
   Icon            =   "Contador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   2070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Acionar tela tipo msn"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Contador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'define o caminho de aviso
    vcaminhodosom = App.Path & "\swintro.wav"
    AcionaAviso ("ohhhh !!!, não precisa de ocx")
End Sub
