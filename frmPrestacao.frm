VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{6A8ED97B-A20C-462C-8D48-6E20F00977AE}#1.0#0"; "RMCALC127.OCX"
Begin VB.Form frmPrestacao 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Controle de Prestações"
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command8 
      Caption         =   "end"
      Height          =   255
      Left            =   10320
      TabIndex        =   23
      Top             =   480
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3615
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Controle de Prestações"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   3075
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Localizar"
      Height          =   465
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Localiza um Registro"
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "B&aixar Lançamento"
      Height          =   465
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Baixar (Pagar) a conta atual"
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Excluir Lançamento"
      Height          =   465
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Exclui a conta Selecionada"
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10560
      TabIndex        =   9
      ToolTipText     =   "Remover Filtro"
      Top             =   4920
      Width           =   390
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmPrestacao.frx":0000
      Left            =   8040
      List            =   "frmPrestacao.frx":0002
      TabIndex        =   8
      Text            =   "Selecione..."
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programa Imobiliária\Dados\Bdimobiliaria.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Prestacao"
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Baixar 
      DataField       =   "Recebidos"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4440
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFF00&
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
      Left            =   240
      TabIndex        =   6
      Top             =   4920
      Width           =   3615
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Código"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   4680
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Vendedor"
      Height          =   195
      Left            =   1440
      TabIndex        =   4
      Top             =   4680
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comprador"
      Height          =   195
      Left            =   2760
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Informações"
      Height          =   465
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "C&alculadora"
      Height          =   465
      Left            =   5880
      TabIndex        =   1
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Editar vencimento"
      Height          =   465
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   1935
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmPrestacao.frx":0004
      Height          =   3135
      Left            =   120
      OleObjectBlob   =   "frmPrestacao.frx":0018
      TabIndex        =   22
      Top             =   840
      Width           =   10935
   End
   Begin VB.Line Line1 
      X1              =   4560
      X2              =   9960
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line2 
      X1              =   4920
      X2              =   10080
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   8040
      TabIndex        =   21
      Top             =   4560
      Width           =   675
   End
   Begin RmCalc.Calc Calc 
      Left            =   5280
      Top             =   2280
      _ExtentX        =   979
      _ExtentY        =   1455
      CorFundo        =   65535
      CorButao        =   32896
      CorFonteTitulo  =   -2147483630
      CorTitulo       =   32896
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Recebido:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   8040
      TabIndex        =   20
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblValor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   8100
      TabIndex        =   19
      Top             =   4320
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Pago:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   9360
      TabIndex        =   18
      Top             =   4080
      Width           =   705
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Left            =   9240
      TabIndex        =   17
      Top             =   4320
      Width           =   840
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   10200
      TabIndex        =   16
      Top             =   4320
      Width           =   840
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comissão:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   225
      Left            =   10320
      TabIndex        =   15
      Top             =   4080
      Width           =   675
   End
End
Attribute VB_Name = "frmPrestacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()

If Combo1.Text = "Vencendo Hoje" Then
    Data1.RecordSource = "SELECT * FROM Vencimentos WHERE Vencimento Like '" & Date & "*'"
    Data1.Refresh

ElseIf Combo1.Text = "Recibos Vencidos" Then
    Data1.RecordSource = "SELECT * FROM Vencimentos WHERE Vencimento < #" & Date & "#"
    Data1.Refresh

ElseIf Combo1.Text = "Recibos Pagos" Then
    Data1.RecordSource = "SELECT * FROM Vencimentos WHERE Recebidos IS NOT NULL"
    Data1.Refresh
    
ElseIf Combo1.Text = "Recibos Não Pagos" Then
    Data1.RecordSource = "SELECT * FROM Vencimentos WHERE Recebidos IS NULL"
    Data1.Refresh
    
ElseIf Combo1.Text = "Recibos com Multa" Then
    Data1.RecordSource = "SELECT * FROM Vencimentos WHERE Multa IS NOT NULL"
    Data1.Refresh
    
ElseIf Combo1.Text = ": REMOVER FILTRO :" Then
    Data1.RecordSource = "SELECT * FROM Vencimentos"
    Data1.Refresh
End If
End Sub

Private Sub Command1_Click()
MsgBox ("Observação: " & Data1.Recordset("Ob1"))
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &H80C0FF
Command1.Font.Bold = True
Command3.BackColor = &H8000000F
Command3.Font.Bold = False
End Sub

Private Sub Command2_Click()
Unload Me
Calc.Abrir
End Sub

Private Sub Command3_Click()
EXCLUIR.Show vbModal, Me
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &H8000000F
Command1.Font.Bold = False
Command3.BackColor = &H80C0FF
Command3.Font.Bold = True
Command4.BackColor = &H8000000F
Command4.Font.Bold = False
Command7.BackColor = &H8000000F
Command7.Font.Bold = False
End Sub

Private Sub Command4_Click()
On Error Resume Next

If frmAluguel.Visible = True Then
If Len(Baixar.Text) > 0 Then
    MsgBox ("Esse lançamento já foi dado baixa em: " & Baixar.Text & " !!"), vbCritical, "Erro ao tentar dar baixa"
ElseIf Len(Baixar.Text) = 0 Then
    frmBaixar.Option1.Value = False
    frmBaixar.Option2.Value = False
    frmBaixar.Text2 = Data1.Recordset.Fields(6)
    frmBaixar.Text6 = Data1.Recordset.Fields(5)
    frmBaixar.txtData.SetFocus
    frmBaixar.Show 1
End If

ElseIf frmPrestacao.Visible = True Then
If Len(Baixar.Text) > 0 Then
    MsgBox ("Esse lançamento já foi dado baixa em: " & Baixar.Text & " !!"), vbCritical, "Erro ao tentar dar baixa"
ElseIf Len(Baixar.Text) = 0 Then
    frmBaixar.Option1.Value = False
    frmBaixar.Option2.Value = False
    frmBaixar.Text2 = Data1.Recordset.Fields(4)
    frmBaixar.txtData.SetFocus
    frmBaixar.Show 1
End If
End If
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.BackColor = &H8000000F
Command3.Font.Bold = False
Command4.BackColor = &H80C0FF
Command4.Font.Bold = True
Command7.BackColor = &H8000000F
Command7.Font.Bold = False
End Sub

Private Sub Command5_Click()
If Option1.Value = True Then
    If Text1.Text = "" Then
        Data1.RecordSource = "SELECT * FROM PRESTACAO"
        Data1.Refresh
        Exit Sub
    End If

    Data1.RecordSource = "SELECT * FROM PRESTACAO WHERE Codigo Like '" & Text1.Text & "*'"
    Data1.Refresh

ElseIf Option2.Value = True Then
    If Text1.Text = "" Then
        Data1.RecordSource = "SELECT * FROM PRESTACAO"
        Data1.Refresh
        Exit Sub
    End If

    Data1.RecordSource = "SELECT * FROM PRESTACAO WHERE VENDEDOR Like '" & Text1.Text & "*'"
    Data1.Refresh
    
ElseIf Option3.Value = True Then
    If Text1.Text = "" Then
        Data1.RecordSource = "SELECT * FROM PRESTACAO"
        Data1.Refresh
        Exit Sub
    End If

    Data1.RecordSource = "SELECT * FROM PRESTACAO WHERE COMPRADOR Like '" & Text1.Text & "*'"
    Data1.Refresh
End If
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command5.BackColor = &H80C0FF
Command5.Font.Bold = True
End Sub

Private Sub Command6_Click()
Data1.RecordSource = "SELECT * FROM PRESTACAO"
Data1.Refresh
End Sub

Private Sub Command7_Click()
On Error Resume Next

frmEditar.Text1 = Data1.Recordset("Vencimento")
frmEditar.Text2 = Data1.Recordset("Recebidos")
frmEditar.Text3 = Data1.Recordset("Valor")
frmEditar.Text4 = Data1.Recordset("Prop")
frmEditar.Text5 = Data1.Recordset("ValorProp")
frmEditar.Text6 = Data1.Recordset("Multa")
frmEditar.Text7 = Data1.Recordset("Iptu")
frmEditar.Show 1

End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &H8000000F
Command1.Font.Bold = False
Command3.BackColor = &H8000000F
Command3.Font.Bold = False
Command4.BackColor = &H8000000F
Command4.Font.Bold = False
Command5.BackColor = &H8000000F
Command5.Font.Bold = False
Command7.BackColor = &H80C0FF
Command7.Font.Bold = True
End Sub

Private Sub Command8_Click()
Unload frmPrestacao
RedefineFormPrincipal
End Sub

Private Sub DBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
On Error GoTo erro
If ColIndex >= 0 And ColIndex <= 9 Then
    Cancel = True
    MsgBox "Não pode ser alterado o conteudo desta célula.", vbCritical, "Aviso!"
    Exit Sub
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Aviso": Exit Sub
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\Dados\Bdimobiliaria.mdb"
Data1.RecordSource = "Prestacao"

With Combo1
    .AddItem "Vencendo Hoje"
    .AddItem "Recibos Vencidos"
    .AddItem "Recibos Pagos"
    .AddItem "Recibos Não Pagos"
    .AddItem "Recibos com Multa"
    .AddItem ": REMOVER FILTRO :"
End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &H8000000F
Command1.Font.Bold = False
Command3.BackColor = &H8000000F
Command3.Font.Bold = False
Command4.BackColor = &H8000000F
Command4.Font.Bold = False
Command5.BackColor = &H8000000F
Command5.Font.Bold = False
Command7.BackColor = &H8000000F
Command7.Font.Bold = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmFundo.Enabled = True
RedefineFormPrincipal
End Sub

Private Sub Option1_Click()
Text1 = ""
Text1.SetFocus
End Sub

Private Sub Option2_Click()
Text1 = ""
Text1.SetFocus
End Sub

Private Sub Option3_Click()
Text1 = ""
Text1.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command5.SetFocus
End If
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
