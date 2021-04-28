VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGeraRecibos 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5280
      TabIndex        =   13
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3480
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Inicia"
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   4440
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   480
      Width           =   6375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Gerar Recibos"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Sair"
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstFaturas 
      Height          =   2775
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4895
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Proprietário"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Locatário"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "dd/mm/aa"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Valor R$"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Contratos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   885
   End
   Begin VB.Label txtTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Prazo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Inicio Contrato:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   5
      Top             =   960
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Término Contrato:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Valor do Recibo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5160
      TabIndex        =   3
      Top             =   960
      Width           =   1440
   End
End
Attribute VB_Name = "frmGeraRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click()
Set Rs3 = Nothing
Rs3.Open "Select Prazo, Inicio, Final, Aluguel From Contrato where Codigo='" & Combo1.Text & "'", Cn, adOpenStatic, adLockPessimistic
CarregaText
End Sub

Private Sub Command3_Click()
Combo1.ListIndex = -1
Text1 = ""
Text2 = ""
Text3 = ""
Text3 = ""
End Sub

Private Sub Form_Load()
Conecta
CarregaCombo
End Sub

Private Function CarregaCombo()
Combo1.Clear
    Do While Not Rs3.EOF
        With Combo1
            .AddItem Rs3!Locador
        End With
Rs3.MoveNext
    Loop
Rs3.Close
End Function

Private Sub Text_Change(Index As Integer)
Select Case Index
End Select
End Sub

Private Function CarregaText()
Text1.Text = Rs3!Prazo
Text2.Text = Rs3!Inicio
Text3.Text = Rs3!Final
Text4.Text = Rs3!Aluguel
End Function
