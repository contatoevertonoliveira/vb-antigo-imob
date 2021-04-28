VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPesq 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Super Imob - Busca Clientes"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Sa&ir"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Verific&ar Dados"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Data dtaDados 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programa Imobiliária\Dados\Bdimobiliaria.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clientes"
      Top             =   4080
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      Caption         =   "Digite para pesquisar:"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "P&esquisar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3120
         Picture         =   "Pesquisar.frx":0000
         TabIndex        =   4
         ToolTipText     =   "Pesquisar"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resultado:"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4455
      Begin MSDBGrid.DBGrid GridBd 
         Bindings        =   "Pesquisar.frx":0B00
         Height          =   2655
         Left            =   120
         OleObjectBlob   =   "Pesquisar.frx":0B17
         TabIndex        =   3
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmPesq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ID As Double

Private Sub Command1_Click()

If Text1.Text = "" Then
    dtaDados.RecordSource = "SELECT * FROM CLIENTES"
    dtaDados.Refresh
    Exit Sub
End If

dtaDados.RecordSource = "SELECT * FROM Clientes WHERE Nome Like '" & Text1.Text & "*'"
dtaDados.Refresh
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
If frmContrLoc.Visible = True Then
    frmClientes.Command1.Enabled = True
End If
    ID = GridBd.Columns(0)
    frmClientes.Show
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
dtaDados.DatabaseName = App.Path & "\dados\Bdimobiliaria.MDB"
dtaDados.RecordSource = "Clientes"
End Sub

Private Sub GridBd_Click()
Command2.Enabled = True
End Sub

Private Sub GridBd_DblClick()
If frmContrLoc.Visible = True Then
    ID = GridBd.Columns(0)
    frmClientes.Show
    frmClientes.Command1.Enabled = True
Else
    ID = GridBd.Columns(0)
    frmClientes.Show
End If
End Sub

Private Sub Text1_Change()

If Len(Text1) = 0 Then
    Command1.Enabled = False
Else
    Command1.Enabled = True
End If

End Sub
