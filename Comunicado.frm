VERSION 5.00
Begin VB.Form frmComunicado 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   1980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Digite o Formulário:"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmComunicado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Text1 = 1 Then
frmContrLoc.ssPainel.Tab = 0
End If

If Text1 = 2 Then
frmContrLoc.ssPainel.Tab = 1
End If

If Text1 = 3 Then
frmContrLoc.ssPainel.Tab = 2
End If

If Text1 = 4 Then
frmContrLoc.ssPainel.Tab = 3
End If

If Text1 = 5 Then
frmContrLoc.ssPainel.Tab = 4
End If

If Text1 > 5 Then
MsgBox ("Você não pode inserir um cliente neste formulário!. Digite outro!")
Text1.SetFocus
End If
End Sub

Private Sub Text1_GotFocus()
Text1.BackColor = &HFFFF&
End Sub

Private Sub Text1_LostFocus()
Text1.BackColor = &H80000005
End Sub
