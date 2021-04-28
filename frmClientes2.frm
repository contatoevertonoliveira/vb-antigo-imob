VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmClientes2 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtBusca 
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
      Height          =   315
      Left            =   120
      TabIndex        =   33
      Top             =   480
      Width           =   2415
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmClientes2.frx":0000
      Height          =   1095
      Left            =   2760
      OleObjectBlob   =   "frmClientes2.frx":0014
      TabIndex        =   32
      Top             =   120
      Width           =   6495
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Nome"
      Height          =   195
      Left            =   1800
      TabIndex        =   31
      Top             =   240
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Código"
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   ">>> B&uscar >>>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      DataField       =   "Codigo"
      DataSource      =   "Data1"
      Enabled         =   0   'False
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
      Left            =   360
      TabIndex        =   17
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Nome"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   360
      MaxLength       =   50
      TabIndex        =   1
      Top             =   2280
      Width           =   6615
   End
   Begin VB.TextBox Text3 
      DataField       =   "Opção"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   7080
      MaxLength       =   50
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      DataField       =   "Endereço"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   360
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2880
      Width           =   3975
   End
   Begin VB.TextBox Text5 
      DataField       =   "Bairro"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4440
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2880
      Width           =   3135
   End
   Begin VB.TextBox Text6 
      DataField       =   "Cidade"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   360
      MaxLength       =   50
      TabIndex        =   6
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox Text7 
      DataField       =   "Fone1"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4560
      MaxLength       =   10
      TabIndex        =   8
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      DataField       =   "Observação"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   4080
      Width           =   8655
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   9135
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAlterar 
         Caption         =   "&Alterar"
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   5760
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   7560
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   9135
      Begin VB.Data Data1 
         Caption         =   "Pessoas"
         Connect         =   "Access"
         DatabaseName    =   "C:\Programa Imobiliária\Dados\Pessoas.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Clientes"
         Top             =   360
         Width           =   7095
      End
      Begin VB.TextBox Text9 
         DataField       =   "Uf"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   7
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox Text10 
         DataField       =   "Cep"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   7560
         MaxLength       =   9
         TabIndex        =   5
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         DataField       =   "Fone2"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6840
         MaxLength       =   10
         TabIndex        =   9
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome da Pessoa / Cliente / Razão Social:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Abreviação:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   6960
         TabIndex        =   27
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   4320
         TabIndex        =   25
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3360
         TabIndex        =   23
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cep:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   7560
         TabIndex        =   22
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone Residencial:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   4440
         TabIndex        =   21
         Top             =   2040
         Width           =   1275
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone Comercial:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   6840
         TabIndex        =   20
         Top             =   2040
         Width           =   1140
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observação:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   2640
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmClientes2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bd As DAO.Database
Dim Tb As DAO.Recordset
Public Sql As String

Private Sub cmdAlterar_Click()

Data1.Recordset.Edit
AbreCaixas
cmdAlterar.Enabled = False
cmdGravar.Enabled = True
cmdNovo.Enabled = False
cmdExcluir.Enabled = False
cmdSair.Caption = "&Cancelar"

End Sub

Private Sub cmdConfirma_Click()

If Option1.Value = True Then
txtBusca = Format(txtBusca, "000")
    If txtBusca.Text = "" Then
        Data1.RecordSource = "SELECT * FROM CLIENTES"
        Data1.Refresh
        Exit Sub
    End If

    Data1.RecordSource = "SELECT * FROM CLIENTES WHERE CODIGO Like '" & txtBusca.Text & "*'"
    Data1.Refresh

ElseIf Option2.Value = True Then
    If txtBusca.Text = "" Then
        Data1.RecordSource = "SELECT * FROM CLIENTES"
        Data1.Refresh
        Exit Sub
    End If

    Data1.RecordSource = "SELECT * FROM CLIENTES WHERE NOME Like '" & txtBusca.Text & "*'"
    Data1.Refresh
End If
End Sub

Private Sub cmdExcluir_Click()
If MsgBox("Confirma Exclusão do Cliente?  -> " & Data1.Recordset![codigo], vbQuestion + vbYesNo, "Excluir Clientes") = vbYes Then
   Data1.Recordset.Delete
   Data1.Refresh
End If
End Sub

Private Sub cmdGravar_Click()

Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
FechaCaixas
cmdGravar.Enabled = False
cmdAlterar.Enabled = True
cmdNovo.Enabled = True
cmdExcluir.Enabled = True
cmdSair.Caption = "&Sair"
MsgBox ("Dados cadastrados com sucesso...")

End Sub

Private Sub cmdNovo_Click()
Dim Novo As String

Set Bd = OpenDatabase(App.Path & "\Dados\Pessoas.MDB")
Set Tb = Bd.OpenRecordset("Clientes", dbOpenTable)

Tb.Index = "Codigo"
Tb.MoveLast
Novo = Tb!codigo
Tb.Seek "=", Novo
If Tb.NoMatch = False Then
    Novo = Novo + 1
End If
    
    Data1.Recordset.AddNew
    AbreCaixas
    limpa
    Text1 = Novo
    Text2.SetFocus
    cmdNovo.Enabled = False
    cmdAlterar.Enabled = False
    cmdGravar.Enabled = True
    cmdExcluir.Enabled = False
    cmdSair.Caption = "&Cancelar"

End Sub

Private Sub cmdSair_Click()
On Error Resume Next
If cmdSair.Caption = "&Cancelar" Then
    Data1.Recordset.CancelUpdate
    Data1.Recordset.MoveLast
    cmdNovo.Enabled = True
    cmdAlterar.Enabled = True
    cmdGravar.Enabled = False
    cmdExcluir.Enabled = True
    FechaCaixas
    cmdSair.Caption = "&Sair"
Else
If MsgBox("Quer sair do Cadastro de Pessoas?", vbYesNo, "Sair do Cadastro") = vbYes Then
    Unload frmClientes2
    RedefineFormPrincipal
  Else
    Exit Sub
End If
End If
frmFundo.Enabled = True
End Sub

Private Sub DBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
On Error GoTo erro
If ColIndex >= 0 And ColIndex <= 3 Then
    Cancel = True
    MsgBox "Não pode ser alterado o conteudo desta célula.", vbCritical, "Aviso!"
    Exit Sub
End If

Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Aviso": Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    SendKeys ("{TAB}")
    KeyAscii = 0
End If
End Sub

Private Sub Form_Load()

Data1.DatabaseName = App.Path & "\dados\pessoas.MDB"
Data1.RecordSource = "Clientes"

End Sub

Private Function FechaCaixas()

Text1.Enabled = False
Text2.Enabled = False
Text2.BackColor = &H8000000A
Text3.Enabled = False
Text3.BackColor = &H8000000A
Text4.Enabled = False
Text4.BackColor = &H8000000A
Text5.Enabled = False
Text5.BackColor = &H8000000A
Text6.Enabled = False
Text6.BackColor = &H8000000A
Text7.Enabled = False
Text7.BackColor = &H8000000A
Text8.Enabled = False
Text8.BackColor = &H8000000A
Text9.Enabled = False
Text9.BackColor = &H8000000A
Text10.Enabled = False
Text10.BackColor = &H8000000A
Text11.Enabled = False
Text11.BackColor = &H8000000A

Label4.Enabled = False
Label5.Enabled = False
Label6.Enabled = False
Label7.Enabled = False
Label8.Enabled = False
Label9.Enabled = False
Label10.Enabled = False
Label11.Enabled = False
Label12.Enabled = False
Label13.Enabled = False

End Function

Private Function AbreCaixas()

Text2.Enabled = True
Text2.BackColor = &H80000005
Text3.Enabled = True
Text3.BackColor = &H80000005
Text4.Enabled = True
Text4.BackColor = &H80000005
Text5.Enabled = True
Text5.BackColor = &H80000005
Text6.Enabled = True
Text6.BackColor = &H80000005
Text7.Enabled = True
Text7.BackColor = &H80000005
Text8.Enabled = True
Text8.BackColor = &H80000005
Text9.Enabled = True
Text9.BackColor = &H80000005
Text10.Enabled = True
Text10.BackColor = &H80000005
Text11.Enabled = True
Text11.BackColor = &H80000005

Label4.Enabled = True
Label5.Enabled = True
Label6.Enabled = True
Label7.Enabled = True
Label8.Enabled = True
Label9.Enabled = True
Label10.Enabled = True
Label11.Enabled = True
Label12.Enabled = True
Label13.Enabled = True

End Function

Private Function limpa()

Text2 = Empty
Text3 = Empty
Text4 = Empty
Text5 = Empty
Text6 = Empty
Text7 = Empty
Text8 = Empty
Text9 = Empty
Text10 = Empty
Text11 = Empty

End Function

Private Sub Form_Unload(Cancel As Integer)
frmFundo.Enabled = True
RedefineFormPrincipal
End Sub

Private Sub Option1_Click()
Data1.RecordSource = "SELECT * FROM Clientes"
Data1.Refresh
txtBusca = ""
txtBusca.SetFocus
End Sub

Private Sub Option2_Click()
Data1.RecordSource = "SELECT * FROM Clientes"
Data1.Refresh
txtBusca = ""
txtBusca.SetFocus
End Sub

Private Sub Text1_Change()
Text1 = Format(Text1.Text, "000")
End Sub

Private Sub Text2_GotFocus()
Text2.BackColor = &HE0E0E0
Text2.Font.Name = "Tahoma"
Text2.Font.Bold = True
End Sub

Private Sub Text2_LostFocus()
Text2 = StrConv(Text2.Text, vbUpperCase)
Text2.BackColor = &H80000005
Text2.Font.Name = "Ms Sans Serif"
Text2.Font.Bold = False
End Sub

Private Sub Text3_GotFocus()
Text3.BackColor = &HE0E0E0
Text3.Font.Name = "Tahoma"
Text3.Font.Bold = True
End Sub

Private Sub Text3_LostFocus()
Text3 = StrConv(Text3.Text, vbUpperCase)
Text3.BackColor = &H80000005
Text3.Font.Name = "Ms Sans Serif"
Text3.Font.Bold = False
End Sub

Private Sub Text4_GotFocus()
Text4.BackColor = &HE0E0E0
Text4.Font.Name = "Tahoma"
Text4.Font.Bold = True
End Sub

Private Sub Text4_LostFocus()
Text4 = StrConv(Text4.Text, vbUpperCase)
Text4.BackColor = &H80000005
Text4.Font.Name = "Ms Sans Serif"
Text4.Font.Bold = False
End Sub

Private Sub Text5_GotFocus()
Text5.BackColor = &HE0E0E0
Text5.Font.Name = "Tahoma"
Text5.Font.Bold = True
End Sub

Private Sub Text5_LostFocus()
Text5 = StrConv(Text5.Text, vbUpperCase)
Text5.BackColor = &H80000005
Text5.Font.Name = "Ms Sans Serif"
Text5.Font.Bold = False
End Sub

Private Sub Text6_GotFocus()
Text6.BackColor = &HE0E0E0
Text6.Font.Name = "Tahoma"
Text6.Font.Bold = True
End Sub

Private Sub Text6_LostFocus()
Text6 = StrConv(Text6.Text, vbUpperCase)
Text6.BackColor = &H80000005
Text6.Font.Name = "Ms Sans Serif"
Text6.Font.Bold = False
End Sub

Private Sub Text7_GotFocus()
Text7.BackColor = &HE0E0E0
Text7.Font.Name = "Tahoma"
Text7.Font.Bold = True
End Sub

Private Sub Text7_LostFocus()
Text7 = StrConv(Text7.Text, vbUpperCase)
Text7.BackColor = &H80000005
Text7.Font.Name = "Ms Sans Serif"
Text7.Font.Bold = False
End Sub

Private Sub Text8_GotFocus()
Text8.BackColor = &HE0E0E0
Text8.Font.Name = "Tahoma"
Text8.Font.Bold = True
End Sub

Private Sub Text8_LostFocus()
Text8 = StrConv(Text8.Text, vbUpperCase)
Text8.BackColor = &H80000005
Text8.Font.Name = "Ms Sans Serif"
Text8.Font.Bold = False
End Sub

Private Sub Text9_GotFocus()
Text9.BackColor = &HE0E0E0
Text9.Font.Name = "Tahoma"
Text9.Font.Bold = True
End Sub

Private Sub Text9_LostFocus()
Text9 = StrConv(Text9.Text, vbUpperCase)
Text9.BackColor = &H80000005
Text9.Font.Name = "Ms Sans Serif"
Text9.Font.Bold = False
End Sub

Private Sub Text10_GotFocus()
Text10.BackColor = &HE0E0E0
Text10.Font.Name = "Tahoma"
Text10.Font.Bold = True
End Sub

Private Sub Text10_LostFocus()
Text10 = StrConv(Text10.Text, vbUpperCase)
Text10.BackColor = &H80000005
Text10.Font.Name = "Ms Sans Serif"
Text10.Font.Bold = False
End Sub

Private Sub Text11_GotFocus()
Text11.BackColor = &HE0E0E0
Text11.Font.Name = "Tahoma"
Text11.Font.Bold = True
End Sub

Private Sub Text11_LostFocus()
Text11 = StrConv(Text11.Text, vbUpperCase)
Text11.BackColor = &H80000005
Text11.Font.Name = "Ms Sans Serif"
Text11.Font.Bold = False
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
