VERSION 5.00
Begin VB.Form frmAlteraDados 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alterar Recibo"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "C&ancelar"
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2880
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Alterar Recibo"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.TextBox Text8 
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   2040
         Width           =   4455
      End
      Begin VB.TextBox Text7 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3720
         TabIndex        =   14
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3960
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox Text2 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Obs:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   330
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Multa:"
         Height          =   195
         Left            =   3120
         TabIndex        =   13
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IPTU:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   195
         Left            =   3480
         TabIndex        =   10
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aluguél de:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Locatário:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Locador:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmAlteraDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

frmRecibos.Label15.ForeColor = &HFF0000
frmRecibos.Label15 = Text1
frmRecibos.Label16.ForeColor = &HFF0000
frmRecibos.Label16 = Text2
frmRecibos.Label17.ForeColor = &HFF0000
frmRecibos.Label17 = Text3
frmRecibos.Label25.ForeColor = &HFF0000
frmRecibos.Label25 = Text5
frmRecibos.Label26.ForeColor = &HFF0000
frmRecibos.Label26 = Text6
frmRecibos.Label27.ForeColor = &HFF0000
frmRecibos.Label27 = Text7

If Text8.Text = "" Then
    frmRecibos.Label12.Caption = "OBS.:"
Else
    frmRecibos.Label12.Visible = True
    frmRecibos.Label12.ForeColor = &HFF0000
    frmRecibos.Label12 = "OBS.: " & Text8
End If

frmRecibos.Label28.ForeColor = &HFF0000

Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

