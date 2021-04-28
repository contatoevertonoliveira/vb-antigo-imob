VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmContas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Números de Instalação"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programa Imobiliária\Dados\Bdimobiliaria.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Contrato"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programa Imobiliária\Dados\Bdimobiliaria.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Numeros"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programa Imobiliária\Dados\Bdimobiliaria.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Contas"
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "B&uscar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   19
      Top             =   4320
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3720
      Width           =   2655
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frmContas.frx":0000
      Height          =   1575
      Left            =   2880
      OleObjectBlob   =   "frmContas.frx":0014
      TabIndex        =   17
      Top             =   3720
      Width           =   5655
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000014&
      Caption         =   "Cadastr&ar Contas"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000014&
      Caption         =   "Cadastr&ar Nº Instalação"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dados Contrato:"
      Height          =   1575
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Width           =   4815
      Begin VB.TextBox Text5 
         BackColor       =   &H0080C0FF&
         DataField       =   "aLocatario"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H0080C0FF&
         DataField       =   "Locador"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H0080C0FF&
         DataField       =   "ID"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Locatário:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Locador:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Números de Instalação:"
      Height          =   1215
      Left            =   3720
      TabIndex        =   3
      Top             =   1800
      Width           =   4815
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         DataSource      =   "Data2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   14
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         DataSource      =   "Data2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Luz (Número de Instalação):"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1995
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Água (Número do Cliente):"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1860
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contratos"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmContas.frx":0EE3
         Height          =   2535
         Left            =   120
         OleObjectBlob   =   "frmContas.frx":0EF7
         TabIndex        =   2
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bd As DAO.Database
Dim Tb As DAO.Recordset

Private Sub Command1_Click()

If Text2.Text = "" Then
    MsgBox ("Digite algum número para gravar!")
    Text2.SetFocus
ElseIf Text6.Text = "" Then
    MsgBox ("Digite algum número para gravar!")
    Text6.SetFocus
Else
    Data2.Recordset.AddNew
    Data2.Recordset("Codigo") = Text3
    Data2.Recordset("Locador") = Text4
    Data2.Recordset("Locatario") = Text5
    Data2.Recordset("Agua") = Text2
    Data2.Recordset("Luz") = Text6
    Data2.UpdateRecord
    MsgBox ("Números de instalação cadastrados com sucesso!")
    Text2 = ""
    Text6 = ""
End If
End Sub

Private Sub Command2_Click()
Data3.Recordset.AddNew
frmConta.Text1 = Text3
frmConta.Text3 = Text5
frmConta.Show 1
End Sub

Private Sub DBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
On Error GoTo erro

If ColIndex >= 0 And ColIndex <= 2 Then
    Cancel = True
    MsgBox "Não pode ser alterado o conteudo desta célula.", vbCritical, "Aviso!"
    Exit Sub
End If

Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Aviso": Exit Sub
End Sub

Private Sub Form_Load()
With Combo1
    .AddItem ": Selecione uma opção :"
    .AddItem ": Código :"
    .AddItem ": Locatário :"
    .AddItem ": Mês de Referência :"
End With
Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
RedefineFormPrincipal
End Sub

Private Sub Text1_Change()
On Error Resume Next

If Text1 = "" Then
Data1.RecordSource = "SELECT * FROM CONTRATO"
Data1.Refresh
End If

Data2.RecordSource = "SELECT * FROM NUMEROS WHERE LOCADOR Like '" & Text1 & "*'"
Data2.Refresh
Text2 = Data2.Recordset("Agua")
Text6 = Data2.Recordset("Luz")

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
    
    If Text1.Text = "" Then
    Data1.RecordSource = "SELECT * FROM CONTRATO"
    Data1.Refresh
    Exit Sub
    End If
End If

Data1.RecordSource = "SELECT * FROM Contrato WHERE Locador Like '" & Text1.Text & "*'"
Data1.Refresh

Data2.RecordSource = "SELECT * FROM NUMEROS WHERE LOCADOR Like '" & Text1 & "*'"
Data2.Refresh
Text2 = Data2.Recordset("Agua")
Text6 = Data2.Recordset("Luz")
End Sub

Private Sub Text2_LostFocus()
Dim NumAgua As String
Dim NumLuz As String

Set Bd = OpenDatabase(App.Path & "\Dados\Bdimobiliaria.MDB")
Set Tb = Bd.OpenRecordset("Numeros", dbOpenTable)
Tb.Index = "Numeros"
Tb.MoveLast
NumAgua = Tb!Agua
NumLuz = Tb!luz
Tb.Seek "=", NumAgua
If Tb.NoMatch = False Then
    MsgBox ("Esse Número de Instalação já está cadastrado!")
    Text2 = ""
    Text6 = ""
    Text2.SetFocus
End If
End Sub

Private Sub Text4_Change()
On Error Resume Next

Data2.RecordSource = "SELECT * FROM NUMEROS WHERE LOCADOR Like '" & Text4 & "*'"
Data2.Refresh
Text2 = Data2.Recordset("Agua")
Text6 = Data2.Recordset("Luz")

End Sub

Private Function RedefineFormPrincipal()

If frmFundo.Text1.Text = "Basico" Then
    frmFundo.cmdCadastros.Enabled = True
    frmFundo.cmdContratos.Enabled = True
    frmFundo.cmdRecibos.Enabled = False
    frmFundo.cmdPrest.Enabled = False
    frmFundo.cmdSair.Caption = "Sair do Programa"
End If
If frmFundo.Text1.Text = "Intermediario" Then
    frmFundo.cmdCadastros.Enabled = True
    frmFundo.cmdContratos.Enabled = True
    frmFundo.cmdRecibos.Enabled = True
    frmFundo.cmdPrest.Enabled = True
    frmFundo.cmdAlugueis.Enabled = True
    frmFundo.cmdSair.Caption = "Sair do Programa"
End If
If frmFundo.Text1.Text = "Avançado" Then
    frmFundo.cmdCadastros.Enabled = True
    frmFundo.cmdContratos.Enabled = True
    frmFundo.cmdRecibos.Enabled = True
    frmFundo.cmdPrest.Enabled = True
    frmFundo.cmdAlugueis.Enabled = True
    frmFundo.cmdSair.Caption = "Sair do Programa"
End If
End Function
