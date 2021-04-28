VERSION 5.00
Object = "{6A8ED97B-A20C-462C-8D48-6E20F00977AE}#1.0#0"; "RMCALC127.OCX"
Begin VB.Form frmFundo 
   BackColor       =   &H8000000C&
   Caption         =   "Super Imob"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmFundo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   6960
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   10320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   1455
      Left            =   6600
      TabIndex        =   15
      Top             =   8040
      Visible         =   0   'False
      Width           =   2655
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Autorização Administração"
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
         Left            =   0
         TabIndex        =   19
         Top             =   1200
         Width           =   2655
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Proposta de Compra"
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
         TabIndex        =   18
         Top             =   840
         Width           =   2175
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Compra e Venda"
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
         Left            =   480
         TabIndex        =   17
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Locação"
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
         Left            =   840
         TabIndex        =   16
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3840
      TabIndex        =   11
      Top             =   8040
      Visible         =   0   'False
      Width           =   2775
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pessoas"
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
         Left            =   840
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Imóveis"
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
         TabIndex        =   13
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clientes"
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
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   1575
      Left            =   3840
      TabIndex        =   3
      Top             =   6360
      Width           =   8175
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair do Programa"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5400
         TabIndex        =   9
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton cmdPrest 
         Caption         =   "Controle Prestações"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5400
         TabIndex        =   8
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdAlugueis 
         Caption         =   "Controle Aluguéis"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   7
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton cmdRecibos 
         Caption         =   "Emissão Recibos"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdContratos 
         Caption         =   "Contratos"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton cmdCadastros 
         Caption         =   "Cadastros"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
   End
   Begin Project1.UserControl1 VBX 
      Left            =   1320
      Top             =   1800
      _ExtentX        =   2355
      _ExtentY        =   2355
   End
   Begin VB.TextBox Text2 
      Height          =   2415
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "frmFundo.frx":0442
      Left            =   8040
      List            =   "frmFundo.frx":0444
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programa Imobiliária\Dados\Bdimobiliaria.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Prestacao"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programa Imobiliária\Dados\Bdimobiliaria.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Vencimentos"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   4080
      Picture         =   "frmFundo.frx":0446
      ScaleHeight     =   4515
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   1560
      Width           =   7575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   960
      Width           =   9015
   End
   Begin RmCalc.Calc Calc 
      Left            =   1000
      Top             =   3000
      _ExtentX        =   979
      _ExtentY        =   1455
      CorFundo        =   12632256
      CorButao        =   192
      CorFonte        =   33023
      CorFonteTitulo  =   128
      CorTitulo       =   255
   End
   Begin VB.Menu mnuImob 
      Caption         =   "&Abrir"
      Enabled         =   0   'False
      Begin VB.Menu itmCon 
         Caption         =   "C&ontratos"
         Begin VB.Menu smnuLoc 
            Caption         =   "Locaç&ão"
            Shortcut        =   {F2}
         End
         Begin VB.Menu smnuCompra 
            Caption         =   "C&ompra e Venda"
            Shortcut        =   {F3}
         End
         Begin VB.Menu smnuProp 
            Caption         =   "&Proposta de Compra"
            Shortcut        =   {F4}
         End
         Begin VB.Menu smnuAd 
            Caption         =   "&Administração"
            Shortcut        =   {F5}
         End
      End
      Begin VB.Menu itmcad 
         Caption         =   "Cad&astro"
         Begin VB.Menu itmcli 
            Caption         =   "&Clientes"
            Shortcut        =   ^C
         End
         Begin VB.Menu itmIm 
            Caption         =   "Im&óveis"
            Shortcut        =   {F8}
         End
         Begin VB.Menu itmPessoas 
            Caption         =   "Pessoas"
            Shortcut        =   {F9}
         End
      End
      Begin VB.Menu itmAlerta 
         Caption         =   "Al&erta"
         Shortcut        =   {F11}
      End
      Begin VB.Menu itmRec 
         Caption         =   "&Emissão de Recibos"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuContr 
         Caption         =   "Controle de Aluguéis"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuPrest 
         Caption         =   "Controle de Prestação"
         Shortcut        =   ^Z
      End
      Begin VB.Menu itmseparador 
         Caption         =   "-"
      End
      Begin VB.Menu itmSair 
         Caption         =   "&Sair"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuOp 
      Caption         =   "Opç&ões"
      Begin VB.Menu itmCalen 
         Caption         =   "Cal&endário"
         Shortcut        =   ^R
      End
      Begin VB.Menu itmcalc 
         Caption         =   "C&alculadora"
      End
   End
End
Attribute VB_Name = "frmFundo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAlugueis_Click()
frmFundo.Enabled = False
frmAluguel.Show
End Sub

Private Sub cmdCadastros_Click()
cmdSair.Caption = "Cancelar"
cmdCadastros.Enabled = False

    If cmdCadastros.Caption = "Cadastros" Then
                
        Frame2.Visible = True
        cmdCadastros.Caption = ""
        cmdContratos.Enabled = False
        cmdRecibos.Enabled = False
        cmdPrest.Enabled = False
        cmdAlugueis.Enabled = False
    
     End If
    
    If cmdCadastros.Caption = "Clientes" Then
        frmClientes.Show
        Frame2.Visible = False
        cmdCadastros.Caption = "Cadastros"
        frmFundo.Enabled = False
    End If
     
    If cmdCadastros.Caption = "Imóveis" Then
        frmImoveis.Show
        Frame2.Visible = False
        cmdCadastros.Caption = "Cadastros"
        frmFundo.Enabled = False
    End If
    
    If cmdCadastros.Caption = "Pessoas" Then
        frmClientes2.Show
        Frame2.Visible = False
        cmdCadastros.Caption = "Cadastros"
        frmFundo.Enabled = False
    End If
    
End Sub

Private Sub cmdContratos_Click()
cmdSair.Caption = "Cancelar"
cmdContratos.Enabled = False

If cmdContratos.Caption = "Contratos" Then
                
        Frame3.Visible = True
        cmdContratos.Caption = ""
        cmdCadastros.Enabled = False
        cmdRecibos.Enabled = False
        cmdPrest.Enabled = False
        cmdAlugueis.Enabled = False
                    
End If
    
    If cmdContratos.Caption = "Locação" Then
        frmContrLoc.Show
        Frame3.Visible = False
        cmdContratos.Caption = "Contratos"
        frmFundo.Enabled = False
    End If
    
    If cmdContratos.Caption = "Compra e Venda" Then
        frmCompraVenda.Show
        Frame3.Visible = False
        cmdContratos.Caption = "Contratos"
        frmFundo.Enabled = False
    End If
     
    If cmdContratos.Caption = "Proposta" Then
        MsgBox ("Desculpe o transtorno!")
        Frame3.Visible = False
        cmdContratos.Caption = "Contratos"
        frmFundo.Enabled = Enabled
    End If
    
    If cmdContratos.Caption = "Autorização" Then
        MsgBox ("Desculpe o transtorno!")
        Frame3.Visible = False
        cmdContratos.Caption = "Contratos"
        frmFundo.Enabled = Enabled
    End If
End Sub

Private Sub cmdPrest_Click()
frmFundo.Enabled = False
frmPrestacao.Show
End Sub

Private Sub cmdRecibos_Click()
frmFundo.Enabled = False
frmRecibos.Show
End Sub

Private Sub cmdSair_Click()

If cmdSair.Caption = "Cancelar" Then
    If Text1.Text = "Basico" Then
        cmdCadastros.Enabled = True
        cmdContratos.Enabled = True
        cmdRecibos.Enabled = False
        cmdPrest.Enabled = False
        cmdAlugueis.Enabled = False
        Frame2.Visible = False
        Frame3.Visible = False
        RedefineBotao
        
    End If
    If Text1.Text = "Intermediario" Then
        cmdCadastros.Enabled = True
        cmdContratos.Enabled = True
        cmdRecibos.Enabled = True
        cmdPrest.Enabled = True
        cmdAlugueis.Enabled = True
        Frame2.Visible = False
        Frame3.Visible = False
        RedefineBotao
    End If
    If Text1.Text = "Avançado" Then
        cmdCadastros.Enabled = True
        cmdContratos.Enabled = True
        cmdRecibos.Enabled = True
        cmdPrest.Enabled = True
        cmdAlugueis.Enabled = True
        Frame2.Visible = False
        Frame3.Visible = False
        RedefineBotao
    End If
    
ElseIf cmdSair.Caption = "Sair do Programa" Then
        If MsgBox("Deseja Sair do Super Imob?", vbYesNo, "Sair do Programa") = vbYes Then
            End
        Else
            Exit Sub
        End If
End If
End Sub

Private Sub Form_Load()
Label1.Caption = "Imobiliária Kielek"
Data1.DatabaseName = App.Path & "\DADOS\BDIMOBILIARIA.MDB"
Data1.RecordSource = "VENCIMENTOS"
Verifica
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub itmcadi_Click()
frmCadastroSenha.Show
End Sub

Private Sub itmcalc_Click()
Calc.Abrir
End Sub

Private Sub itmCalen_Click()
frmCalendario.Show 1
End Sub

Private Sub itmcli_Click()
frmClientes.Show 1
End Sub

Private Sub itmIm_Click()
frmImoveis.Show 1
End Sub

Private Sub itmPesq_Click()
frmOpContrato.Show
End Sub

Private Sub itmPessoas_Click()
frmClientes2.Show 1
End Sub

Private Sub itmRec_Click()
frmRecibos.Show 1
End Sub

Private Sub itmSair_Click()
End
End Sub

Private Sub mnuContr_Click()
frmAluguel.Show 1
End Sub

Private Sub mnuPrest_Click()
frmPrestacao.Show 1
End Sub

Private Sub Option1_Click()
cmdCadastros.Caption = "Clientes"
cmdCadastros.Enabled = True
End Sub

Private Sub Option2_Click()
cmdCadastros.Caption = "Imóveis"
cmdCadastros.Enabled = True
End Sub

Private Sub Option3_Click()
cmdCadastros.Caption = "Pessoas"
cmdCadastros.Enabled = True
End Sub

Private Sub Option4_Click()
cmdContratos.Caption = "Locação"
cmdContratos.Enabled = True
End Sub

Private Sub Option5_Click()
cmdContratos.Caption = "Compra e Venda"
cmdContratos.Enabled = True
End Sub

Private Sub Option6_Click()
cmdContratos.Caption = "Proposta"
cmdContratos.Enabled = True
End Sub

Private Sub Option7_Click()
cmdContratos.Caption = "Autorização"
cmdContratos.Enabled = True
End Sub

Private Sub smnuCompra_Click()
frmCompraVenda.Show
End Sub

Private Sub smnuLoc_Click()
frmContrLoc.Show 1
End Sub

Private Sub smnurec_Click()
frmRecibos.Show
End Sub

Private Function Verifica()
On Error Resume Next
Dim Informacao As String
Dim list As Variant
Dim strSentence As String
Dim counter As Integer

Data1.RecordSource = "SELECT * FROM Vencimentos WHERE Vencimento Like '" & Text1.Text & "*'"
Data1.Refresh

Data2.RecordSource = "SELECT * FROM Prestacao WHERE Vencimento Like '" & CDate(Text1.Text) & "*'"
Data2.Refresh

If Data1.Recordset("Vencimento") = "" Then

    List1.AddItem "Até o momento está tudo ok, continue trabalhando!"
    Text2 = "Até o momento está tudo ok, continue trabalhando!"
    
    If InStr(Text2.Text, vbNewLine) Then
    list = Split(Text2.Text, vbNewLine)
    
    For counter = 0 To UBound(list)
    strSentence = list(counter)
        VBX.Parse strSentence
    Next counter
    Else
    VBX.Parse Text2.Text
    End If
    
ElseIf Data1.Recordset("Vencimento") = Text1.Text Then

    List1.AddItem "Existem aluguéis vencendo hoje!, Confira!"
    Text2 = "Existem aluguéis vencendo hoje!, Confira!"
    
    If InStr(Text2.Text, vbNewLine) Then
    list = Split(Text2.Text, vbNewLine)
    
    For counter = 0 To UBound(list)
    strSentence = list(counter)
        VBX.Parse strSentence
    Next counter
    Else
    VBX.Parse Text2.Text
    End If
End If
End Function

Private Function RedefineBotao()

    cmdCadastros.Caption = "Cadastros"
    cmdContratos.Caption = "Contratos"
    cmdRecibos.Caption = "Emissão Recibos"
    cmdAlugueis.Caption = "Controle Aluguéis"
    cmdPrest.Caption = "Controle Prestações"
    cmdSair.Caption = "Sair do Programa"
    
    Option1.Value = False
    Option2.Value = False
    Option3.Value = False
    Option4.Value = False
    Option5.Value = False
    Option6.Value = False
    Option7.Value = False

End Function
