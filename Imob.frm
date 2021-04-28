VERSION 5.00
Begin VB.Form frmShow 
   BorderStyle     =   0  'None
   Caption         =   "Super Imob"
   ClientHeight    =   4470
   ClientLeft      =   2535
   ClientTop       =   2355
   ClientWidth     =   7470
   Icon            =   "Imob.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Imob.frx":0442
   ScaleHeight     =   4470
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Carregando...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      TabIndex        =   0
      Top             =   3960
      Width           =   2280
   End
End
Attribute VB_Name = "frmShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Top = (Screen.Height - Height) / 2
Me.Left = (Screen.Width - Width) / 2
MousePointer = 12
End Sub

Private Sub Timer1_Timer()
Load frmFundo
Unload Me
frmFundo.Show
frmFundo.Enabled = False
Frmlogin.Show
Frmlogin.Enabled = True
End Sub
