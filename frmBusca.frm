VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBusca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Super Imob - Busca"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "frmBusca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6615
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data dtaBusca 
      Caption         =   "busca"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "<< V&er dados >>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame4 
      Caption         =   "Resultado Consulta:"
      Height          =   2295
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   6135
      Begin MSFlexGridLib.MSFlexGrid MGrid1 
         Height          =   1935
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   25
         FixedCols       =   0
         AllowBigSelection=   -1  'True
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "Codigo  |  Locador                                         |  Locatário                                  "
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Inici&ar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Digite algum nome para buscar:"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   6135
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   5895
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Selecione a forma de consulta:"
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   6135
      Begin VB.ComboBox cboTable 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   5895
      End
      Begin VB.ComboBox cboOp 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   5895
      End
   End
   Begin VB.CommandButton cmdsair 
      Caption         =   "&Sair da Pesquisa"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(""Dê um duplo clique para ir até o registro selecionado"")"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1320
      TabIndex        =   9
      Top             =   4680
      Width           =   3945
   End
End
Attribute VB_Name = "frmBusca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bd As DAO.Database
Dim Rs As DAO.Recordset
Dim vCodigo As String
Dim vNome As String
Dim vNac As String
Dim vProf As String
Dim vCivil As String
Dim vRg As String
Dim vCpf As String
Dim vEnd As String
Dim VNum As String
Dim vBr As String
Dim vCid As String
Dim vUf As String
Dim vCep As String
Dim vTelRes As String
Dim vTelCom As String
Dim vEmail As String
Dim vSite As String
Dim vConjuge As String
Dim vNacConj As String
Dim vProfConj As String
Dim vRgConj As String
Dim vCpfConj As String
Dim txt0 As String
Dim txt1 As String
Dim txt2 As String
Dim txt3 As String
Dim txt4 As String
Dim txt5 As String
Dim txt6 As String
Dim txt7 As String
Dim txt8 As String
Dim txt9 As String
Dim txt10 As String
Dim txt11 As String
Dim txt12 As String
Dim txt13 As String
Dim txt14 As String
Dim txt15 As String
Dim txt16 As String
Dim txt17 As String
Dim txt18 As String
Dim txt19 As String
Dim txt20 As String
Dim txt21 As String
Dim txt22 As String
Dim txt23 As String
Dim txt24 As String
Dim txt25 As String
Dim txt26 As String
Dim txt27 As String
Dim txt28 As String
Dim txt29 As String
Dim txt30 As String
Dim txt31 As String
Dim txt32 As String
Dim txt33 As String
Dim txt34 As String
Dim txt35 As String
Dim txt36 As String
Dim txt37 As String
Dim txt38 As String
Dim txt39 As String
Dim txt40 As String
Dim txt41 As String
Dim txt42 As String
Dim txt43 As String
Dim txt44 As String
Dim txt45 As String
Dim txt46 As String
Dim txt47 As String
Dim txt48 As String
Dim txt49 As String
Dim txt50 As String
Dim txt51 As String
Dim txt52 As String
Dim txt53 As String
Dim txt54 As String
Dim txt55 As String
Dim txt56 As String
Dim txt57 As String
Dim txt58 As String
Dim txt59 As String
Dim txt60 As String
Dim txt61 As String
Dim txt62 As String
Dim txt63 As String
Dim txt64 As String
Dim txt65 As String
Dim txt66 As String
Dim txt67 As String
Dim txt68 As String
Dim txt69 As String
Dim txt70 As String
Dim txt71 As String
Dim txt72 As String
Dim txt73 As String
Dim txt74 As String
Dim txt75 As String
Dim txt76 As String
Dim txt77 As String
Dim txt78 As String
Dim txt79 As String
Dim txt80 As String
Dim txt81 As String
Dim txt82 As String
Dim txt83 As String
Dim txt84 As String
Dim txt85 As String
Dim txt86 As String
Dim txt87 As String
Dim txt88 As String
Dim txt89 As String
Dim txt90 As String
Dim txt91 As String

Private Sub cboOp_Click()
Text1.Enabled = True
Command2.Visible = False
Text1 = ""
Text1.SetFocus
End Sub

Private Sub cboTable_Click()
If cboTable.Text = "Contratos / Locação" Then
    Text1 = ""
    cboOp.Clear
    cboOp.AddItem "Codigo"
    cboOp.AddItem "Locador"
    MGrid1.Cols = 92
    MGrid1.Clear
    ConfigMgrid1
End If

If cboTable.Text = "Contratos / Compra e Venda Novos" Then
    Text = ""
    cboOp.Clear
    cboOp.AddItem "Codigo"
    cboOp.AddItem "Vendedor"
    MGrid2.Cols = 36
    MGrid2.Clear
End If

If cboTable.Text = "Contratos / Compra e Venda Antigos" Then
    Text = ""
    cboOp.Clear
    cboOp.AddItem "Codigo"
    cboOp.AddItem "Vendedor"
    MGrid2.Cols = 3
    MGrid2.Clear
End If

If cboTable.Text = "Clientes" Then
    Text = ""
    cboOp.Clear
    cboOp.AddItem "Codigo"
    cboOp.AddItem "Nome"
    MGrid1.Cols = 25
    MGrid1.Clear
    ConfigMgrid
End If
cboOp.Enabled = True
Command1.Enabled = True
Command2.Visible = False
End Sub

Private Sub cmdSair_Click()

If MsgBox("Quer sair da Busca?", vbYesNo, "Sair da Busca") = vbYes Then
    Unload frmBusca
Else
    Exit Sub
End If
End Sub

Private Sub Command1_Click()

Command1.Enabled = False
Command2.Visible = False
Text1 = ""
Text1.Enabled = False
cboOp.ListIndex = -1
cboTable.ListIndex = -1
cboTable.SetFocus

End Sub

Private Sub Command2_Click()

If Text1.Text = "" Then
    MsgBox ("Digite algum nome para a pesquisa!")
    Text1.SetFocus

Else

    If frmContrLoc.Visible = True Then
        If cboTable = "Clientes" Then
            If cboOp = "Nome" Then
                Verifica
                frmContrLoc.Label33.Caption = "Clientes"
                frmContrLoc.Label34.Caption = "Nome"
            End If
        Else
            If cboOp = "Codigo" Then
                Verifica
                frmContrLoc.Label33.Caption = "Clientes"
                frmContrLoc.Label34.Caption = "Nome"
            End If
            End If
        Else
            If cboTable = "Clientes" Then
                Unload Me
                frmClientes.Show
                LimparClientes
                CarregaClientes
            End If
        End If
End If

If frmContrLoc.Visible = True Then
        If cboTable = "Contratos / Locação" Then
            If cboOp = "Codigo" Then
                Contrato
                TxtFechado
            End If
         Else
            If cboOp = "Locador" Then
                Contrato
                TxtFechado
                frmContrLoc.Show
            End If
         End If
End If

If frmContrLoc.Visible = False Then
        If cboTable = "Contratos / Locação" Then
            If cboOp = "Codigo" Then
                frmContrLoc.Show
                Verifica
                TxtFechado
            End If
         Else
            If cboOp = "Locador" Then
                frmContrLoc.Show
                Verifica
                TxtFechado
            End If
         End If
End If
End Sub

Private Sub Form_Load()

Connection

With cboOp
        .AddItem "Codigo"
        .AddItem "Nome"
End With
With cboTable
        .AddItem "Clientes"
        .AddItem "Contratos / Locação"
        .AddItem "Contratos / Compra e Venda Novos"
        .AddItem "Contratos / Compra e Venda Antigos"
End With
ConfigMgrid
End Sub

Private Sub MGrid1_Click()
Dim Posit As Single
Dim Marcador As String

If cboTable = "Clientes" Then
Posit = MGrid1.Row

txt0 = MGrid1.TextMatrix(Posit, 0)
txt1 = MGrid1.TextMatrix(Posit, 1)
txt2 = MGrid1.TextMatrix(Posit, 2)
txt3 = MGrid1.TextMatrix(Posit, 3)
txt4 = MGrid1.TextMatrix(Posit, 4)
txt5 = MGrid1.TextMatrix(Posit, 5)
txt6 = MGrid1.TextMatrix(Posit, 6)
txt7 = MGrid1.TextMatrix(Posit, 7)
txt8 = MGrid1.TextMatrix(Posit, 8)
txt9 = MGrid1.TextMatrix(Posit, 9)
txt10 = MGrid1.TextMatrix(Posit, 10)
txt11 = MGrid1.TextMatrix(Posit, 11)
txt12 = MGrid1.TextMatrix(Posit, 12)
txt13 = MGrid1.TextMatrix(Posit, 13)
txt14 = MGrid1.TextMatrix(Posit, 14)
txt15 = MGrid1.TextMatrix(Posit, 15)
txt16 = MGrid1.TextMatrix(Posit, 16)
txt17 = MGrid1.TextMatrix(Posit, 17)
txt18 = MGrid1.TextMatrix(Posit, 18)
txt19 = MGrid1.TextMatrix(Posit, 19)
txt20 = MGrid1.TextMatrix(Posit, 20)
txt21 = MGrid1.TextMatrix(Posit, 21)
txt22 = MGrid1.TextMatrix(Posit, 22)
txt23 = MGrid1.TextMatrix(Posit, 23)
txt24 = MGrid1.TextMatrix(Posit, 24)
End If

If cboTable = "Contratos / Locação" Then
Posit = MGrid1.Row

txt0 = MGrid1.TextMatrix(Posit, 0)
txt1 = MGrid1.TextMatrix(Posit, 1)
txt2 = MGrid1.TextMatrix(Posit, 2)
txt3 = MGrid1.TextMatrix(Posit, 3)
txt4 = MGrid1.TextMatrix(Posit, 4)
txt5 = MGrid1.TextMatrix(Posit, 5)
txt6 = MGrid1.TextMatrix(Posit, 6)
txt7 = MGrid1.TextMatrix(Posit, 7)
txt8 = MGrid1.TextMatrix(Posit, 8)
txt9 = MGrid1.TextMatrix(Posit, 9)
txt10 = MGrid1.TextMatrix(Posit, 10)
txt11 = MGrid1.TextMatrix(Posit, 11)
txt12 = MGrid1.TextMatrix(Posit, 12)
txt13 = MGrid1.TextMatrix(Posit, 13)
txt14 = MGrid1.TextMatrix(Posit, 14)
txt15 = MGrid1.TextMatrix(Posit, 15)
txt16 = MGrid1.TextMatrix(Posit, 16)
txt17 = MGrid1.TextMatrix(Posit, 17)
txt18 = MGrid1.TextMatrix(Posit, 18)
txt19 = MGrid1.TextMatrix(Posit, 19)
txt20 = MGrid1.TextMatrix(Posit, 20)
txt21 = MGrid1.TextMatrix(Posit, 21)
txt22 = MGrid1.TextMatrix(Posit, 22)
txt23 = MGrid1.TextMatrix(Posit, 23)
txt24 = MGrid1.TextMatrix(Posit, 24)
txt25 = MGrid1.TextMatrix(Posit, 25)
txt26 = MGrid1.TextMatrix(Posit, 26)
txt27 = MGrid1.TextMatrix(Posit, 27)
txt28 = MGrid1.TextMatrix(Posit, 28)
txt29 = MGrid1.TextMatrix(Posit, 29)
txt30 = MGrid1.TextMatrix(Posit, 30)
txt31 = MGrid1.TextMatrix(Posit, 31)
txt32 = MGrid1.TextMatrix(Posit, 32)
txt33 = MGrid1.TextMatrix(Posit, 33)
txt34 = MGrid1.TextMatrix(Posit, 34)
txt35 = MGrid1.TextMatrix(Posit, 35)
txt36 = MGrid1.TextMatrix(Posit, 36)
txt37 = MGrid1.TextMatrix(Posit, 37)
txt38 = MGrid1.TextMatrix(Posit, 38)
txt39 = MGrid1.TextMatrix(Posit, 39)
txt40 = MGrid1.TextMatrix(Posit, 40)
txt41 = MGrid1.TextMatrix(Posit, 41)
txt42 = MGrid1.TextMatrix(Posit, 42)
txt43 = MGrid1.TextMatrix(Posit, 43)
txt44 = MGrid1.TextMatrix(Posit, 44)
txt45 = MGrid1.TextMatrix(Posit, 45)
txt46 = MGrid1.TextMatrix(Posit, 46)
txt47 = MGrid1.TextMatrix(Posit, 47)
txt48 = MGrid1.TextMatrix(Posit, 48)
txt49 = MGrid1.TextMatrix(Posit, 49)
txt50 = MGrid1.TextMatrix(Posit, 50)
txt51 = MGrid1.TextMatrix(Posit, 51)
txt52 = MGrid1.TextMatrix(Posit, 52)
txt53 = MGrid1.TextMatrix(Posit, 53)
txt54 = MGrid1.TextMatrix(Posit, 54)
txt55 = MGrid1.TextMatrix(Posit, 55)
txt56 = MGrid1.TextMatrix(Posit, 56)
txt57 = MGrid1.TextMatrix(Posit, 57)
txt58 = MGrid1.TextMatrix(Posit, 58)
txt59 = MGrid1.TextMatrix(Posit, 59)
txt60 = MGrid1.TextMatrix(Posit, 60)
txt61 = MGrid1.TextMatrix(Posit, 61)
txt62 = MGrid1.TextMatrix(Posit, 62)
txt63 = MGrid1.TextMatrix(Posit, 63)
txt64 = MGrid1.TextMatrix(Posit, 64)
txt65 = MGrid1.TextMatrix(Posit, 65)
txt66 = MGrid1.TextMatrix(Posit, 66)
txt67 = MGrid1.TextMatrix(Posit, 67)
txt68 = MGrid1.TextMatrix(Posit, 68)
txt69 = MGrid1.TextMatrix(Posit, 69)
txt70 = MGrid1.TextMatrix(Posit, 70)
txt71 = MGrid1.TextMatrix(Posit, 71)
txt72 = MGrid1.TextMatrix(Posit, 72)
txt73 = MGrid1.TextMatrix(Posit, 73)
txt74 = MGrid1.TextMatrix(Posit, 74)
txt75 = MGrid1.TextMatrix(Posit, 75)
txt76 = MGrid1.TextMatrix(Posit, 76)
txt77 = MGrid1.TextMatrix(Posit, 77)
txt78 = MGrid1.TextMatrix(Posit, 78)
txt79 = MGrid1.TextMatrix(Posit, 79)
txt80 = MGrid1.TextMatrix(Posit, 80)
txt81 = MGrid1.TextMatrix(Posit, 81)
txt82 = MGrid1.TextMatrix(Posit, 82)
txt83 = MGrid1.TextMatrix(Posit, 83)
txt84 = MGrid1.TextMatrix(Posit, 84)
txt85 = MGrid1.TextMatrix(Posit, 85)
txt86 = MGrid1.TextMatrix(Posit, 86)
txt87 = MGrid1.TextMatrix(Posit, 87)
txt88 = MGrid1.TextMatrix(Posit, 88)
txt89 = MGrid1.TextMatrix(Posit, 89)
txt90 = MGrid1.TextMatrix(Posit, 90)
txt91 = MGrid1.TextMatrix(Posit, 91)
End If

    If cboTable = "Clientes" Then
        dtaBusca.DatabaseName = "\Dados\Bdimobiliaria.mdb"
        dtaBusca.RecordSource = "Loc"
        dtaBusca.RecordsetType = 0
    
    dtaBusca.Recordset.Index = "IndCod"
    dtaBusca.Recordset.Seek "=", txt0
    If dtaBusca.Recordset.NoMatch = True Then
        MsgBox "Cliente não localizado ! ", vbExclamation, "Localizar Clientes"
    End If
    dtaBusca.Recordset.MovePrevious
    Marcador = dtaBusca.Recordset.Bookmark
    frmContrLoc.dtaContrato.Recordset.Bookmark = Marcador
    End If
If Text1.Text = "" Then
Command2.Visible = False
Else
Command2.Visible = True
End If
End Sub

Private Sub MGrid1_DblClick()

If Text1.Text = "" Then
    MsgBox ("Digite algum nome para a pesquisa!")
    Text1.SetFocus

Else

    If frmContrLoc.Visible = True Then
        If cboTable = "Clientes" Then
            If cboOp = "Nome" Then
                Verifica
                frmContrLoc.Label33.Caption = "Clientes"
                frmContrLoc.Label34.Caption = "Nome"
            End If
        Else
            If cboOp = "Codigo" Then
                Verifica
                frmContrLoc.Label33.Caption = "Clientes"
                frmContrLoc.Label34.Caption = "Nome"
            End If
            End If
        Else
            If cboTable = "Clientes" Then
                Unload Me
                frmClientes.Show
                LimparClientes
                CarregaClientes
            End If
        End If
End If

If frmContrLoc.Visible = True Then
        If cboTable = "Contratos / Locação" Then
            If cboOp = "Codigo" Then
                Verifica
                TxtFechado
            End If
         Else
            If cboOp = "Locador" Then
                Verifica
                TxtFechado
            End If
         End If
End If

If frmContrLoc.Visible = False Then
        If cboTable = "Contratos / Locação" Then
            If cboOp = "Codigo" Then
                frmContrLoc.Show
                Verifica
                TxtFechado
            End If
         Else
            If cboOp = "Locador" Then
                frmContrLoc.Show
                Verifica
                TxtFechado
            End If
         End If
End If
End Sub

Private Sub Text1_Change()

If cboTable = "Clientes" Then
    If cboOp = "Nome" Then
    
        If Text1.Text = "" Then
            MGrid1.Enabled = False
            vCodigo = ""
            vNome = ""
            vNac = ""
            vProf = ""
            vCivil = ""
            vRg = ""
            vCpf = ""
            vEnd = ""
            VNum = ""
            vBr = ""
            vCid = ""
            vUf = ""
            vCep = ""
            vTelRes = ""
            vTelCom = ""
            vEmail = ""
            vSite = ""
            vConjuge = ""
            vNacConj = ""
            vProfConj = ""
            vRgConj = ""
            vCpfConj = ""
        Else
            MGrid1.Enabled = True
        End If

        If Text1.Text = "" Then
            MGrid1.Rows = 2
            MgridVazio1
            MGrid1.Rows = MGrid1.Rows - 1
            Me.Caption = "Buscar Cliente"
        Exit Sub
        End If

        MGrid1.Rows = 2
        CarregaFlex1
    End If

End If

If cboTable = "Clientes" Then
    If cboOp = "Codigo" Then
    
        If Text1.Text = "" Then
            MGrid1.Enabled = False
            vCodigo = ""
            vNome = ""
            vNac = ""
            vProf = ""
            vCivil = ""
            vRg = ""
            vCpf = ""
            vEnd = ""
            VNum = ""
            vBr = ""
            vCid = ""
            vUf = ""
            vCep = ""
            vTelRes = ""
            vTelCom = ""
            vEmail = ""
            vSite = ""
            vConjuge = ""
            vNacConj = ""
            vProfConj = ""
            vRgConj = ""
            vCpfConj = ""
        Else
            MGrid1.Enabled = True
        End If

    If Text1.Text = "" Then
        MGrid1.Rows = 2
        MgridVazio1
        MGrid1.Rows = MGrid1.Rows - 1
        Me.Caption = "Buscar Cliente"
    Exit Sub
    End If

    MGrid1.Rows = 2
    CarregaFlex3
End If
End If


If cboTable = "Contratos / Locação" Then
    If cboOp = "Locador" Then
    
    If Text1.Text = "" Then
    MGrid1.Enabled = False
    vCodigo = ""
    vNome = ""
    vNac = ""
    vProf = ""
    vCivil = ""
    vRg = ""
    vCpf = ""
    vEnd = ""
    VNum = ""
    vBr = ""
    vCid = ""
    vUf = ""
    vCep = ""
    vTelRes = ""
    vTelCom = ""
    vEmail = ""
    vSite = ""
    vConjuge = ""
    vNacConj = ""
    vProfConj = ""
    vRgConj = ""
    vCpfConj = ""
Else
    MGrid1.Enabled = True
End If

    If Text1.Text = "" Then
    MGrid1.Rows = 2
    MgridVazio
    MGrid1.Rows = MGrid1.Rows - 1
    Me.Caption = "Buscar Cliente"
    Exit Sub
    End If

    MGrid1.Rows = 2
    CarregaFlex
End If
End If


If cboTable = "Contratos / Locação" Then
    If cboOp = "Codigo" Then
    
    If Text1.Text = "" Then
    MGrid1.Enabled = False
    vCodigo = ""
    vNome = ""
    vNac = ""
    vProf = ""
    vCivil = ""
    vRg = ""
    vCpf = ""
    vEnd = ""
    VNum = ""
    vBr = ""
    vCid = ""
    vUf = ""
    vCep = ""
    vTelRes = ""
    vTelCom = ""
    vEmail = ""
    vSite = ""
    vConjuge = ""
    vNacConj = ""
    vProfConj = ""
    vRgConj = ""
    vCpfConj = ""
Else
    MGrid1.Enabled = True
End If

    If Text1.Text = "" Then
    MGrid1.Rows = 2
    MgridVazio
    MGrid1.Rows = MGrid1.Rows - 1
    Me.Caption = "Buscar Cliente"
    Exit Sub
    End If

    MGrid1.Rows = 2
    CarregaFlex2
End If
End If
End Sub

Private Sub Text1_Click()
Command2.Visible = False
End Sub

Private Sub Text1_GotFocus()
Text1.BackColor = &HFFFF&
End Sub

Private Sub Text1_LostFocus()
Text1.BackColor = &H80000005
End Sub


Private Function ConfigMgrid()

MGrid1.ColWidth(0) = 600
MGrid1.ColWidth(1) = 2000
MGrid1.ColWidth(2) = 1100
MGrid1.ColWidth(3) = 1700
MGrid1.ColWidth(4) = 1100
MGrid1.ColWidth(5) = 1200
MGrid1.ColWidth(6) = 1200
MGrid1.ColWidth(7) = 1700
MGrid1.ColWidth(8) = 1700
MGrid1.ColWidth(9) = 1700
MGrid1.ColWidth(10) = 800
MGrid1.ColWidth(11) = 800
MGrid1.ColWidth(12) = 800
MGrid1.ColWidth(13) = 1700
MGrid1.ColWidth(14) = 1200
MGrid1.ColWidth(15) = 800
MGrid1.ColWidth(16) = 1200
MGrid1.ColWidth(17) = 1200
MGrid1.ColWidth(18) = 2000
MGrid1.ColWidth(19) = 2000
MGrid1.ColWidth(20) = 2000
MGrid1.ColWidth(21) = 1100
MGrid1.ColWidth(22) = 1700
MGrid1.ColWidth(23) = 1200
MGrid1.ColWidth(24) = 1200

MGrid1.TextMatrix(0, 0) = "Código"
MGrid1.TextMatrix(0, 1) = "Cliente"
MGrid1.TextMatrix(0, 2) = "Nac"
MGrid1.TextMatrix(0, 3) = "Prof"
MGrid1.TextMatrix(0, 4) = "Civil"
MGrid1.TextMatrix(0, 5) = "Rg"
MGrid1.TextMatrix(0, 6) = "Cpf"
MGrid1.TextMatrix(0, 7) = "End"
MGrid1.TextMatrix(0, 8) = "Bairro"
MGrid1.TextMatrix(0, 9) = "Comp"
MGrid1.TextMatrix(0, 10) = "Num"
MGrid1.TextMatrix(0, 11) = "Ap"
MGrid1.TextMatrix(0, 12) = "Bloc"
MGrid1.TextMatrix(0, 13) = "Cidade"
MGrid1.TextMatrix(0, 14) = "Cep"
MGrid1.TextMatrix(0, 15) = "Uf"
MGrid1.TextMatrix(0, 16) = "1-Tel"
MGrid1.TextMatrix(0, 17) = "2-Tel"
MGrid1.TextMatrix(0, 18) = "E-mail"
MGrid1.TextMatrix(0, 19) = "Site"
MGrid1.TextMatrix(0, 20) = "Conjuge"
MGrid1.TextMatrix(0, 21) = "NacConj"
MGrid1.TextMatrix(0, 22) = "ProfConj"
MGrid1.TextMatrix(0, 23) = "RgConj"
MGrid1.TextMatrix(0, 24) = "CpfConj"

End Function

Private Function ConfigMgrid1()

MGrid1.ColWidth(0) = 600
MGrid1.ColWidth(1) = 2000
MGrid1.ColWidth(2) = 1100
MGrid1.ColWidth(3) = 1700
MGrid1.ColWidth(4) = 1100
MGrid1.ColWidth(5) = 1200
MGrid1.ColWidth(6) = 1200
MGrid1.ColWidth(7) = 1700
MGrid1.ColWidth(8) = 1700
MGrid1.ColWidth(9) = 1700
MGrid1.ColWidth(10) = 800
MGrid1.ColWidth(11) = 800
MGrid1.ColWidth(12) = 2000
MGrid1.ColWidth(13) = 1100
MGrid1.ColWidth(14) = 1700
MGrid1.ColWidth(15) = 1200
MGrid1.ColWidth(16) = 1200

MGrid1.ColWidth(17) = 2000
MGrid1.ColWidth(18) = 1100
MGrid1.ColWidth(19) = 1700
MGrid1.ColWidth(20) = 1100
MGrid1.ColWidth(21) = 1200
MGrid1.ColWidth(22) = 1200
MGrid1.ColWidth(23) = 1700
MGrid1.ColWidth(24) = 1700
MGrid1.ColWidth(25) = 1700
MGrid1.ColWidth(26) = 800
MGrid1.ColWidth(27) = 800
MGrid1.ColWidth(28) = 2000
MGrid1.ColWidth(29) = 1100
MGrid1.ColWidth(30) = 1700
MGrid1.ColWidth(31) = 1200
MGrid1.ColWidth(32) = 1200

MGrid1.ColWidth(33) = 2000
MGrid1.ColWidth(34) = 1100
MGrid1.ColWidth(35) = 1700
MGrid1.ColWidth(36) = 1100
MGrid1.ColWidth(37) = 1200
MGrid1.ColWidth(38) = 1200
MGrid1.ColWidth(39) = 1700
MGrid1.ColWidth(40) = 1700
MGrid1.ColWidth(41) = 1700
MGrid1.ColWidth(42) = 800
MGrid1.ColWidth(43) = 800
MGrid1.ColWidth(44) = 2000
MGrid1.ColWidth(45) = 1100
MGrid1.ColWidth(46) = 1700
MGrid1.ColWidth(47) = 1200
MGrid1.ColWidth(48) = 1200

MGrid1.ColWidth(49) = 2000
MGrid1.ColWidth(50) = 1100
MGrid1.ColWidth(51) = 1700
MGrid1.ColWidth(52) = 1100
MGrid1.ColWidth(53) = 1200
MGrid1.ColWidth(54) = 1200
MGrid1.ColWidth(55) = 1700
MGrid1.ColWidth(56) = 1700
MGrid1.ColWidth(57) = 1700
MGrid1.ColWidth(58) = 800
MGrid1.ColWidth(59) = 800
MGrid1.ColWidth(60) = 2000
MGrid1.ColWidth(61) = 1100
MGrid1.ColWidth(62) = 1700
MGrid1.ColWidth(63) = 1200
MGrid1.ColWidth(64) = 1200

MGrid1.ColWidth(65) = 2000
MGrid1.ColWidth(66) = 1100
MGrid1.ColWidth(67) = 1700
MGrid1.ColWidth(68) = 1100
MGrid1.ColWidth(69) = 1200
MGrid1.ColWidth(70) = 1200
MGrid1.ColWidth(71) = 1700
MGrid1.ColWidth(72) = 1700
MGrid1.ColWidth(73) = 1700
MGrid1.ColWidth(74) = 800
MGrid1.ColWidth(75) = 800
MGrid1.ColWidth(76) = 2000
MGrid1.ColWidth(77) = 1100
MGrid1.ColWidth(78) = 1700
MGrid1.ColWidth(79) = 1200
MGrid1.ColWidth(80) = 1200

MGrid1.ColWidth(81) = 1700
MGrid1.ColWidth(82) = 1700
MGrid1.ColWidth(83) = 1700
MGrid1.ColWidth(84) = 800
MGrid1.ColWidth(85) = 400
MGrid1.ColWidth(86) = 800
MGrid1.ColWidth(87) = 800
MGrid1.ColWidth(88) = 1200
MGrid1.ColWidth(90) = 1200
MGrid1.ColWidth(91) = 2000

MGrid1.TextMatrix(0, 0) = "Código"
MGrid1.TextMatrix(0, 1) = "Locador"
MGrid1.TextMatrix(0, 2) = "Nac"
MGrid1.TextMatrix(0, 3) = "Prof"
MGrid1.TextMatrix(0, 4) = "Civil"
MGrid1.TextMatrix(0, 5) = "Rg"
MGrid1.TextMatrix(0, 6) = "Cpf"
MGrid1.TextMatrix(0, 7) = "End"
MGrid1.TextMatrix(0, 8) = "Bairro"
MGrid1.TextMatrix(0, 9) = "Cidade"
MGrid1.TextMatrix(0, 10) = "Uf"
MGrid1.TextMatrix(0, 11) = "Cep"
MGrid1.TextMatrix(0, 12) = "Conjuge"
MGrid1.TextMatrix(0, 13) = "NacConj"
MGrid1.TextMatrix(0, 14) = "ProfConj"
MGrid1.TextMatrix(0, 15) = "RgConj"
MGrid1.TextMatrix(0, 16) = "CpfConj"

MGrid1.TextMatrix(0, 17) = "Locatário"
MGrid1.TextMatrix(0, 18) = "Nac"
MGrid1.TextMatrix(0, 19) = "Prof"
MGrid1.TextMatrix(0, 20) = "Civil"
MGrid1.TextMatrix(0, 21) = "Rg"
MGrid1.TextMatrix(0, 22) = "Cpf"
MGrid1.TextMatrix(0, 23) = "End"
MGrid1.TextMatrix(0, 24) = "Bairro"
MGrid1.TextMatrix(0, 25) = "Cidade"
MGrid1.TextMatrix(0, 26) = "Uf"
MGrid1.TextMatrix(0, 27) = "Cep"
MGrid1.TextMatrix(0, 28) = "Conjuge"
MGrid1.TextMatrix(0, 29) = "NacConj"
MGrid1.TextMatrix(0, 30) = "ProfConj"
MGrid1.TextMatrix(0, 31) = "RgConj"
MGrid1.TextMatrix(0, 32) = "CpfConj"

MGrid1.TextMatrix(0, 33) = "Locatário"
MGrid1.TextMatrix(0, 34) = "Nac"
MGrid1.TextMatrix(0, 35) = "Prof"
MGrid1.TextMatrix(0, 36) = "Civil"
MGrid1.TextMatrix(0, 37) = "Rg"
MGrid1.TextMatrix(0, 38) = "Cpf"
MGrid1.TextMatrix(0, 39) = "End"
MGrid1.TextMatrix(0, 40) = "Bairro"
MGrid1.TextMatrix(0, 41) = "Cidade"
MGrid1.TextMatrix(0, 42) = "Uf"
MGrid1.TextMatrix(0, 43) = "Cep"
MGrid1.TextMatrix(0, 44) = "Conjuge"
MGrid1.TextMatrix(0, 45) = "NacConj"
MGrid1.TextMatrix(0, 46) = "ProfConj"
MGrid1.TextMatrix(0, 47) = "RgConj"
MGrid1.TextMatrix(0, 48) = "CpfConj"

MGrid1.TextMatrix(0, 49) = "Fiador"
MGrid1.TextMatrix(0, 50) = "Nac"
MGrid1.TextMatrix(0, 51) = "Prof"
MGrid1.TextMatrix(0, 52) = "Civil"
MGrid1.TextMatrix(0, 53) = "Rg"
MGrid1.TextMatrix(0, 54) = "Cpf"
MGrid1.TextMatrix(0, 55) = "End"
MGrid1.TextMatrix(0, 56) = "Bairro"
MGrid1.TextMatrix(0, 57) = "Uf"
MGrid1.TextMatrix(0, 58) = "Cidade"
MGrid1.TextMatrix(0, 59) = "Cep"
MGrid1.TextMatrix(0, 60) = "Conjuge"
MGrid1.TextMatrix(0, 61) = "NacConj"
MGrid1.TextMatrix(0, 62) = "ProfConj"
MGrid1.TextMatrix(0, 63) = "RgConj"
MGrid1.TextMatrix(0, 64) = "CpfConj"

MGrid1.TextMatrix(0, 65) = "Fiador"
MGrid1.TextMatrix(0, 66) = "Nac"
MGrid1.TextMatrix(0, 67) = "Prof"
MGrid1.TextMatrix(0, 68) = "Civil"
MGrid1.TextMatrix(0, 69) = "Rg"
MGrid1.TextMatrix(0, 70) = "Cpf"
MGrid1.TextMatrix(0, 71) = "End"
MGrid1.TextMatrix(0, 72) = "Bairro"
MGrid1.TextMatrix(0, 73) = "Cidade"
MGrid1.TextMatrix(0, 74) = "Uf"
MGrid1.TextMatrix(0, 75) = "Cep"
MGrid1.TextMatrix(0, 76) = "Conjuge"
MGrid1.TextMatrix(0, 77) = "NacConj"
MGrid1.TextMatrix(0, 78) = "ProfConj"
MGrid1.TextMatrix(0, 79) = "RgConj"
MGrid1.TextMatrix(0, 80) = "CpfConj"

MGrid1.TextMatrix(0, 81) = "EndImóvel"
MGrid1.TextMatrix(0, 82) = "Finalidade"
MGrid1.TextMatrix(0, 83) = "BImóvel"
MGrid1.TextMatrix(0, 84) = "Aluguél"
MGrid1.TextMatrix(0, 85) = "Dia"
MGrid1.TextMatrix(0, 86) = "Índice"
MGrid1.TextMatrix(0, 87) = "Período"
MGrid1.TextMatrix(0, 88) = "Prazo"
MGrid1.TextMatrix(0, 89) = "Início"
MGrid1.TextMatrix(0, 90) = "Final"
MGrid1.TextMatrix(0, 91) = "Obs"

End Function

Private Function ConfigMgrid2()

MGrid1.ColWidth(0) = 600
MGrid1.ColWidth(1) = 2000
MGrid1.ColWidth(2) = 1100
MGrid1.ColWidth(3) = 1700
MGrid1.ColWidth(4) = 1100
MGrid1.ColWidth(5) = 1200
MGrid1.ColWidth(6) = 1200
MGrid1.ColWidth(7) = 1700
MGrid1.ColWidth(8) = 1700
MGrid1.ColWidth(9) = 1700
MGrid1.ColWidth(10) = 800
MGrid1.ColWidth(11) = 800
MGrid1.ColWidth(12) = 800
MGrid1.ColWidth(13) = 1700
MGrid1.ColWidth(14) = 1200
MGrid1.ColWidth(15) = 800
MGrid1.ColWidth(16) = 1200
MGrid1.ColWidth(17) = 1200
MGrid1.ColWidth(18) = 2000
MGrid1.ColWidth(19) = 2000
MGrid1.ColWidth(20) = 2000
MGrid1.ColWidth(21) = 1100
MGrid1.ColWidth(22) = 1700
MGrid1.ColWidth(23) = 1200
MGrid1.ColWidth(24) = 1200
MGrid1.ColWidth(25) = 2000
MGrid1.ColWidth(26) = 2000
MGrid1.ColWidth(27) = 1100
MGrid1.ColWidth(28) = 1700
MGrid1.ColWidth(29) = 1200
MGrid1.ColWidth(30) = 1200
MGrid1.ColWidth(31) = 2000
MGrid1.ColWidth(32) = 1100
MGrid1.ColWidth(33) = 1700
MGrid1.ColWidth(34) = 1200
MGrid1.ColWidth(35) = 1200

MGrid1.TextMatrix(0, 0) = "Código"
MGrid1.TextMatrix(0, 1) = "Vendedor"
MGrid1.TextMatrix(0, 2) = "Nac"
MGrid1.TextMatrix(0, 3) = "Prof"
MGrid1.TextMatrix(0, 4) = "Civil"
MGrid1.TextMatrix(0, 5) = "Rg"
MGrid1.TextMatrix(0, 6) = "Cpf"
MGrid1.TextMatrix(0, 7) = "End"
MGrid1.TextMatrix(0, 8) = "Bairro"
MGrid1.TextMatrix(0, 9) = "Comp"
MGrid1.TextMatrix(0, 10) = "Cidade"
MGrid1.TextMatrix(0, 11) = "Cep"
MGrid1.TextMatrix(0, 12) = "Uf"
MGrid1.TextMatrix(0, 13) = "1-Tel"
MGrid1.TextMatrix(0, 14) = "2-Tel"
MGrid1.TextMatrix(0, 15) = "E-mail"
MGrid1.TextMatrix(0, 16) = "Site"
MGrid1.TextMatrix(0, 17) = "Conjuge"
MGrid1.TextMatrix(0, 18) = "NacConj"
MGrid1.TextMatrix(0, 19) = "ProfConj"
MGrid1.TextMatrix(0, 20) = "RgConj"
MGrid1.TextMatrix(0, 21) = "CpfConj"
MGrid1.TextMatrix(0, 22) = "Cidade"
MGrid1.TextMatrix(0, 26) = "Cep"
MGrid1.TextMatrix(0, 27) = "Uf"
MGrid1.TextMatrix(0, 28) = "1-Tel"
MGrid1.TextMatrix(0, 29) = "2-Tel"
MGrid1.TextMatrix(0, 30) = "E-mail"
MGrid1.TextMatrix(0, 31) = "Site"
MGrid1.TextMatrix(0, 32) = "Conjuge"
MGrid1.TextMatrix(0, 33) = "NacConj"
MGrid1.TextMatrix(0, 34) = "ProfConj"
MGrid1.TextMatrix(0, 35) = "RgConj"
MGrid1.TextMatrix(0, 36) = "CpfConj"

End Function

Private Function Color()
Dim p_objeto As Object

For Each p_objeto In frmContrLoc.Controls
    If TypeOf p_objeto Is TextBox Then
    p_objeto.Enabled = True
    p_objeto.BackColor = &HC0FFFF
End If
Next p_objeto
frmContrLoc.Combo1(0).BackColor = &HC0FFFF

End Function

Private Function Tab0()

frmContrLoc.Text3 = txt1
frmContrLoc.Text4 = txt2
frmContrLoc.Text5 = txt3
frmContrLoc.Combo2 = txt4
frmContrLoc.Text6 = txt5
frmContrLoc.Text7 = txt6
frmContrLoc.Text8 = txt7 & ", Nº " & txt10 & " " & "APTO: " & txt11 & " " & "BLOCO: " & txt12
frmContrLoc.Text9 = txt8
frmContrLoc.Text10 = txt13
frmContrLoc.Text11 = txt15
frmContrLoc.Text12 = txt14
frmContrLoc.Text13 = txt20
frmContrLoc.Text14 = txt21
frmContrLoc.Text15 = txt22
frmContrLoc.Text16 = txt23
frmContrLoc.Text17 = txt24
frmContrLoc.Command3.Enabled = False
frmContrLoc.Command4.Enabled = True
frmContrLoc.Command5.Enabled = False
frmContrLoc.Command6.Enabled = False
frmContrLoc.Command9.Enabled = False
frmContrLoc.Label14.Caption = "Dados do Locador"

End Function

Private Function Tab1()

frmContrLoc.Text18 = txt1
frmContrLoc.Text19 = txt2
frmContrLoc.Text20 = txt3
frmContrLoc.Combo3 = txt4
frmContrLoc.Text21 = txt5
frmContrLoc.Text22 = txt6
frmContrLoc.Text23 = txt7 & ", Nº " & txt10 & " " & "APTO: " & txt11 & " " & "BLOCO: " & txt12
frmContrLoc.Text24 = txt8
frmContrLoc.Text25 = txt13
frmContrLoc.Text26 = txt15
frmContrLoc.Text27 = txt14
frmContrLoc.Text28 = txt20
frmContrLoc.Text29 = txt21
frmContrLoc.Text30 = txt22
frmContrLoc.Text31 = txt23
frmContrLoc.Text32 = txt24
frmContrLoc.Command3.Enabled = False
frmContrLoc.Command4.Enabled = True
frmContrLoc.Command5.Enabled = False
frmContrLoc.Command6.Enabled = False
frmContrLoc.Command9.Enabled = False
frmContrLoc.Label14.Caption = "Dados do Locatário"
frmContrLoc.ssPainel.Tab = 1

End Function

Private Function Tab2()

frmContrLoc.Text33 = txt1
frmContrLoc.Text34 = txt2
frmContrLoc.Text35 = txt3
frmContrLoc.Combo4 = txt4
frmContrLoc.Text36 = txt5
frmContrLoc.Text37 = txt6
frmContrLoc.Text38 = txt7 & ", Nº " & txt10 & " " & "APTO: " & txt11 & " " & "BLOCO: " & txt12
frmContrLoc.Text39 = txt8
frmContrLoc.Text40 = txt13
frmContrLoc.Text41 = txt15
frmContrLoc.Text42 = txt14
frmContrLoc.Text43 = txt20
frmContrLoc.Text44 = txt21
frmContrLoc.Text45 = txt22
frmContrLoc.Text46 = txt23
frmContrLoc.Text47 = txt24
frmContrLoc.Command3.Enabled = False
frmContrLoc.Command4.Enabled = True
frmContrLoc.Command5.Enabled = False
frmContrLoc.Command6.Enabled = False
frmContrLoc.Command9.Enabled = False
frmContrLoc.Label14.Caption = "Dados do Locatário"
frmContrLoc.ssPainel.Tab = 2

End Function

Private Function Tab3()

frmContrLoc.Text48 = txt1
frmContrLoc.Text49 = txt2
frmContrLoc.Text50 = txt3
frmContrLoc.Combo5 = txt4
frmContrLoc.Text51 = txt5
frmContrLoc.Text52 = txt6
frmContrLoc.Text53 = txt7 & ", Nº " & txt10 & " " & "APTO: " & txt11 & " " & "BLOCO: " & txt12
frmContrLoc.Text54 = txt8
frmContrLoc.Text55 = txt13
frmContrLoc.Text56 = txt15
frmContrLoc.Text57 = txt14
frmContrLoc.Text58 = txt20
frmContrLoc.Text59 = txt21
frmContrLoc.Text60 = txt22
frmContrLoc.Text61 = txt23
frmContrLoc.Text62 = txt24
frmContrLoc.Command3.Enabled = False
frmContrLoc.Command4.Enabled = True
frmContrLoc.Command5.Enabled = False
frmContrLoc.Command6.Enabled = False
frmContrLoc.Command9.Enabled = False
frmContrLoc.Label14.Caption = "Dados do Fiador"
frmContrLoc.ssPainel.Tab = 3

End Function

Private Function Tab4()

frmContrLoc.Text63 = txt1
frmContrLoc.Text64 = txt2
frmContrLoc.Text65 = txt3
frmContrLoc.Combo1 = txt4
frmContrLoc.Text66 = txt5
frmContrLoc.Text67 = txt6
frmContrLoc.Text68 = txt7 & ", Nº " & txt10 & " " & "APTO: " & txt11 & " " & "BLOCO: " & txt12
frmContrLoc.Text69 = txt8
frmContrLoc.Text70 = txt13
frmContrLoc.Text71 = txt15
frmContrLoc.Text72 = txt14
frmContrLoc.Text73 = txt20
frmContrLoc.Text74 = txt21
frmContrLoc.Text75 = txt22
frmContrLoc.Text76 = txt23
frmContrLoc.Text77 = txt24
frmContrLoc.Command3.Enabled = False
frmContrLoc.Command4.Enabled = True
frmContrLoc.Command5.Enabled = False
frmContrLoc.Command6.Enabled = False
frmContrLoc.Command9.Enabled = False
frmContrLoc.Label14.Caption = "Dados do Fiador"
frmContrLoc.ssPainel.Tab = 4

End Function

Private Function Clientes()
frmClientes.txtcod = txt0
frmClientes.txtCodigo = "KL" & Format(txt0, "000")
frmClientes.txtlocatario = txt1
frmClientes.txtnacional = txt2
frmClientes.TxtProfissao = txt3
frmClientes.CobEstadoCivil = txt4
frmClientes.txtruanot = txt7
frmClientes.txtbairronotifi = txt8
frmClientes.txtcomplenot = txt9
frmClientes.txtnnot = txt10
frmClientes.txtnumAP = txt11
frmClientes.txtbloco = txt12
frmClientes.txtcidadenotifi = txt13
frmClientes.txtcepnotifi = txt14
frmClientes.txtest = txt15
frmClientes.txttelefores = txt16
frmClientes.txttelefonecom = txt17
frmClientes.txtRg = txt5
frmClientes.txtCpf = txt6
frmClientes.txtemail = txt18
frmClientes.txtsite = txt19
frmClientes.txtconjuge = txt20
frmClientes.txtconjugeNacional = txt21
frmClientes.txtProfissaoConjuge = txt22
frmClientes.txtRgconjuge = txt23
frmClientes.txtCpfconjuge = txt24
frmClientes.CmdIncluir.Enabled = False
frmClientes.cmdCancelar.Enabled = False
frmClientes.txtconjuge.Enabled = False
frmClientes.txtconjuge.BackColor = &H8000000B
frmClientes.txtconjugeNacional.Enabled = False
frmClientes.txtconjugeNacional.BackColor = &H8000000B
frmClientes.txtCpfconjuge.Enabled = False
frmClientes.txtCpfconjuge.BackColor = &H8000000B
frmClientes.txtRgconjuge.Enabled = False
frmClientes.txtRgconjuge.BackColor = &H8000000B
frmClientes.txtProfissaoConjuge.Enabled = False
frmClientes.txtProfissaoConjuge.BackColor = &H8000000B
End Function


Private Function Verifica()

        
If frmContrLoc.Visible = True Then
    Unload Me
    Contrato
Else
    If frmContrLoc.ssPainel.Tab = 0 Then
        Unload Me
        Tab0
        Color
        frmContrLoc.Text3.SetFocus
        frmContrLoc.Command1.Enabled = True
        frmContrLoc.Text1.Enabled = False
        frmContrLoc.Text3.Enabled = False
        frmContrLoc.ssPainel.Tab = 0
        frmContrLoc.Show
    End If
    If frmContrLoc.ssPainel.Tab = 1 Then
        Unload Me
        Tab1
        Color
        frmContrLoc.Text18.SetFocus
        frmContrLoc.Command1.Enabled = True
        frmContrLoc.Text1.Enabled = False
        frmContrLoc.Text3.Enabled = False
        frmContrLoc.ssPainel.Tab = 1
        frmContrLoc.Show
    End If
        If frmContrLoc.ssPainel.Tab = 2 Then
            Unload Me
            Tab2
            Color
            frmContrLoc.Text33.SetFocus
            frmContrLoc.Command1.Enabled = True
            frmContrLoc.Text1.Enabled = False
            frmContrLoc.Text3.Enabled = False
            frmContrLoc.ssPainel.Tab = 2
            frmContrLoc.Show
        End If
        If frmContrLoc.ssPainel.Tab = 3 Then
            Unload Me
            Tab3
            Color
            frmContrLoc.Text48.SetFocus
            frmContrLoc.Command1.Enabled = True
            frmContrLoc.Text1.Enabled = False
            frmContrLoc.Text3.Enabled = False
            frmContrLoc.ssPainel.Tab = 3
            frmContrLoc.Show
        End If
        If frmContrLoc.ssPainel.Tab = 4 Then
            Unload Me
            Tab4
            Color
            frmContrLoc.Text63.SetFocus
            frmContrLoc.Command1.Enabled = True
            frmContrLoc.Text1.Enabled = False
            frmContrLoc.Text3.Enabled = False
            frmContrLoc.ssPainel.Tab = 4
            frmContrLoc.Show
        End If
        
        If frmContrLoc.ssPainel.Tab = 5 Then
            MsgBox ("Você não pode inserir um cliente neste formulário!")
        End If
End If
End Function

Private Function Contrato()
Set Bd = OpenDatabase(App.Path & "\Dados\Bdimobiliaria.MDB")
Set Rs = Bd.OpenRecordset("Contrato", dbOpenTable)
Rs.Index = "IndCod"

    Rs.Seek "=", txt0
    If Rs.NoMatch = True Then
    MsgBox ("Código não cadastrado!")
    Rs.MovePrevious
    End If
    CarregaContrato
    frmContrLoc.ssPainel.Tab = 0
    frmContrLoc.Command6.Enabled = False
    frmContrLoc.Command7.Enabled = False

Bd.Close
End Function

Private Function CarregaClientes()
 
dtaBusca.DatabaseName = ""
dtaBusca.DatabaseName = App.Path & "\dados\bdimobiliaria.mdb"
dtaBusca.RecordSource = "Loc"
dtaBusca.RecordsetType = 0
   
    dtaBusca.Recordset.Index = "IndCod"
 
      dtaBusca.Recordset.Seek "=", txt0
      If dtaBusca.Recordset.NoMatch Then
        MsgBox "Cliente não localizado ! ", vbExclamation, "Localizar Clientes"
        dtaBusca.Recordset.Bookmark = Marcador
      End If
        frmContrLoc.dtaContrato.Recordset.Bookmark = Marcador
    
End Function

Private Function TxtFechado()
Dim TxtBox As Object
Dim Opts  As Object

For Each TxtBox In frmContrLoc.Controls
    If TypeOf TxtBox Is TextBox Then
    TxtBox.Enabled = False
    TxtBox.BackColor = &H8000000B
End If
Next TxtBox

For Each Opts In frmContrLoc.Controls
    If TypeOf Opts Is OptionButton Then
    Opts.Enabled = False
End If
Next Opts

End Function

Private Function LimparCaixas()
Dim TxtBox As Object

For Each TxtBox In frmContrLoc.Controls
    If TypeOf TxtBox Is TextBox Then
    TxtBox.Text = ""
End If
Next TxtBox

End Function

Private Function LimparClientes()
Dim TxtBox As Object

For Each TxtBox In frmClientes.Controls
    If TypeOf TxtBox Is TextBox Then
    TxtBox.Text = ""
End If
Next TxtBox

End Function

Private Function CarregaContrato()
Dim criterio As Long
Dim Marcador As Variant
 
dtaBusca.DatabaseName = ""
dtaBusca.DatabaseName = App.Path & "\dados\bdimobiliaria.mdb"
dtaBusca.RecordSource = "Contrato"
dtaBusca.RecordsetType = 0
        
    dtaBusca.Recordset.Index = "indcod"
  
    criterio = InputBox$("Codigo do cliente a localizar: ", "Localizar Clientes")
 
    If criterio <> Empty Then
      dtaBusca.Recordset.Seek "=", criterio
      Marcador = dtaBusca.Recordset.Bookmark
      If dtaBusca.Recordset.NoMatch Then
        MsgBox "Cliente não localizado ! ", vbExclamation, "Localizar Clientes"
        dtaBusca.Recordset.Bookmark = Marcador
      End If
    Else
        dtaBusca.Recordset.Bookmark = Marcador
        frmContrLoc.dtaContrato.Recordset.Bookmark = dtaBusca.Recordset.Bookmark
    End If

     
End Function


Private Function MgridVazio()
MGrid1.TextMatrix(MGrid1.Rows - 1, 0) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 1) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 2) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 3) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 4) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 5) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 6) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 7) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 8) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 9) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 10) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 11) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 12) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 13) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 14) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 15) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 16) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 17) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 18) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 19) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 20) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 21) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 22) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 23) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 24) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 25) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 26) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 27) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 28) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 29) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 30) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 31) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 32) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 33) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 34) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 35) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 36) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 37) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 38) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 39) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 40) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 41) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 42) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 43) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 44) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 45) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 46) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 47) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 48) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 49) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 50) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 51) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 52) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 53) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 54) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 55) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 56) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 57) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 58) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 59) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 60) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 61) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 62) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 63) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 64) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 65) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 66) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 67) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 68) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 69) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 70) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 71) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 72) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 73) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 74) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 75) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 76) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 77) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 78) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 79) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 80) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 81) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 82) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 83) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 84) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 85) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 86) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 87) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 88) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 89) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 90) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 91) = ""
End Function

Private Function MgridVazio1()
MGrid1.TextMatrix(MGrid1.Rows - 1, 0) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 1) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 2) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 3) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 4) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 5) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 6) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 7) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 8) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 9) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 10) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 11) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 12) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 13) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 14) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 15) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 16) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 17) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 18) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 19) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 20) = ""
    MGrid1.TextMatrix(MGrid1.Rows - 1, 21) = ""

End Function

Private Function CarregaFlex()

MGrid1.Rows = 2

Connection

    RSTB.Open "SELECT * FROM Contrato WHERE Locador Like '%" & Text1.Text & "%'", CONE, adOpenStatic, adLockOptimistic

    Do While Not RSTB.EOF
    If RSTB.Fields(0).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 0) = RSTB.Fields(0).Value
    If RSTB.Fields(1).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 1) = RSTB.Fields(1).Value
    If RSTB.Fields(2).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 2) = RSTB.Fields(2).Value
    If RSTB.Fields(3).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 3) = RSTB.Fields(3).Value
    If RSTB.Fields(4).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 4) = RSTB.Fields(4).Value
    If RSTB.Fields(5).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 5) = RSTB.Fields(5).Value
    If RSTB.Fields(6).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 6) = RSTB.Fields(6).Value
    If RSTB.Fields(7).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 7) = RSTB.Fields(7).Value
    If RSTB.Fields(8).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 8) = RSTB.Fields(8).Value
    If RSTB.Fields(9).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 9) = RSTB.Fields(9).Value
    If RSTB.Fields(10).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 10) = RSTB.Fields(10).Value
    If RSTB.Fields(11).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 11) = RSTB.Fields(11).Value
    If RSTB.Fields(12).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 12) = RSTB.Fields(12).Value
    If RSTB.Fields(13).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 13) = RSTB.Fields(13).Value
    If RSTB.Fields(14).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 14) = RSTB.Fields(14).Value
    If RSTB.Fields(15).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 15) = RSTB.Fields(15).Value
    If RSTB.Fields(16).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 16) = RSTB.Fields(16).Value
    If RSTB.Fields(17).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 17) = RSTB.Fields(17).Value
    If RSTB.Fields(18).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 18) = RSTB.Fields(18).Value
    If RSTB.Fields(19).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 19) = RSTB.Fields(19).Value
    If RSTB.Fields(20).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 20) = RSTB.Fields(20).Value
    If RSTB.Fields(21).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 21) = RSTB.Fields(21).Value
    If RSTB.Fields(22).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 22) = RSTB.Fields(22).Value
    If RSTB.Fields(23).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 23) = RSTB.Fields(23).Value
    If RSTB.Fields(24).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 24) = RSTB.Fields(24).Value
    If RSTB.Fields(25).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 25) = RSTB.Fields(25).Value
    If RSTB.Fields(26).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 26) = RSTB.Fields(26).Value
    If RSTB.Fields(27).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 27) = RSTB.Fields(27).Value
    If RSTB.Fields(28).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 28) = RSTB.Fields(28).Value
    If RSTB.Fields(29).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 29) = RSTB.Fields(29).Value
    If RSTB.Fields(30).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 30) = RSTB.Fields(30).Value
    If RSTB.Fields(31).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 31) = RSTB.Fields(31).Value
    If RSTB.Fields(32).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 32) = RSTB.Fields(32).Value
    If RSTB.Fields(33).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 33) = RSTB.Fields(33).Value
    If RSTB.Fields(34).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 34) = RSTB.Fields(34).Value
    If RSTB.Fields(35).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 35) = RSTB.Fields(35).Value
    If RSTB.Fields(36).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 36) = RSTB.Fields(36).Value
    If RSTB.Fields(37).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 37) = RSTB.Fields(37).Value
    If RSTB.Fields(38).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 38) = RSTB.Fields(38).Value
    If RSTB.Fields(39).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 39) = RSTB.Fields(39).Value
    If RSTB.Fields(40).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 40) = RSTB.Fields(40).Value
    If RSTB.Fields(41).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 41) = RSTB.Fields(41).Value
    If RSTB.Fields(42).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 42) = RSTB.Fields(42).Value
    If RSTB.Fields(43).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 43) = RSTB.Fields(43).Value
    If RSTB.Fields(44).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 44) = RSTB.Fields(44).Value
    If RSTB.Fields(45).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 45) = RSTB.Fields(45).Value
    If RSTB.Fields(46).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 46) = RSTB.Fields(46).Value
    If RSTB.Fields(47).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 47) = RSTB.Fields(47).Value
    If RSTB.Fields(48).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 48) = RSTB.Fields(48).Value
    If RSTB.Fields(49).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 49) = RSTB.Fields(49).Value
    If RSTB.Fields(50).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 50) = RSTB.Fields(50).Value
    If RSTB.Fields(51).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 51) = RSTB.Fields(51).Value
    If RSTB.Fields(52).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 52) = RSTB.Fields(52).Value
    If RSTB.Fields(53).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 53) = RSTB.Fields(53).Value
    If RSTB.Fields(54).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 54) = RSTB.Fields(54).Value
    If RSTB.Fields(55).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 55) = RSTB.Fields(55).Value
    If RSTB.Fields(56).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 56) = RSTB.Fields(56).Value
    If RSTB.Fields(57).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 57) = RSTB.Fields(57).Value
    If RSTB.Fields(58).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 58) = RSTB.Fields(58).Value
    If RSTB.Fields(59).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 59) = RSTB.Fields(59).Value
    If RSTB.Fields(60).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 60) = RSTB.Fields(60).Value
    If RSTB.Fields(61).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 61) = RSTB.Fields(61).Value
    If RSTB.Fields(62).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 62) = RSTB.Fields(62).Value
    If RSTB.Fields(63).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 63) = RSTB.Fields(63).Value
    If RSTB.Fields(64).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 64) = RSTB.Fields(64).Value
    If RSTB.Fields(65).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 65) = RSTB.Fields(65).Value
    If RSTB.Fields(66).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 66) = RSTB.Fields(66).Value
    If RSTB.Fields(67).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 67) = RSTB.Fields(67).Value
    If RSTB.Fields(68).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 68) = RSTB.Fields(68).Value
    If RSTB.Fields(69).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 69) = RSTB.Fields(69).Value
    If RSTB.Fields(70).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 70) = RSTB.Fields(70).Value
    If RSTB.Fields(71).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 71) = RSTB.Fields(71).Value
    If RSTB.Fields(72).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 72) = RSTB.Fields(72).Value
    If RSTB.Fields(73).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 73) = RSTB.Fields(73).Value
    If RSTB.Fields(74).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 74) = RSTB.Fields(74).Value
    If RSTB.Fields(75).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 75) = RSTB.Fields(75).Value
    If RSTB.Fields(76).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 76) = RSTB.Fields(76).Value
    If RSTB.Fields(77).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 77) = RSTB.Fields(77).Value
    If RSTB.Fields(78).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 78) = RSTB.Fields(78).Value
    If RSTB.Fields(79).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 79) = RSTB.Fields(79).Value
    If RSTB.Fields(80).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 80) = RSTB.Fields(80).Value
    If RSTB.Fields(81).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 81) = RSTB.Fields(81).Value
    If RSTB.Fields(82).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 82) = RSTB.Fields(82).Value
    If RSTB.Fields(83).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 83) = RSTB.Fields(83).Value
    If RSTB.Fields(84).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 84) = RSTB.Fields(84).Value
    If RSTB.Fields(85).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 85) = RSTB.Fields(85).Value
    If RSTB.Fields(86).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 86) = RSTB.Fields(86).Value
    If RSTB.Fields(87).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 87) = RSTB.Fields(87).Value
    If RSTB.Fields(88).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 88) = RSTB.Fields(88).Value
    If RSTB.Fields(89).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 89) = RSTB.Fields(89).Value
    If RSTB.Fields(90).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 90) = RSTB.Fields(90).Value
    If RSTB.Fields(91).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 91) = RSTB.Fields(91).Value
    
    MGrid1.Rows = MGrid1.Rows + 1
    RSTB.MoveNext
    
    Loop

    MGrid1.Rows = MGrid1.Rows - 1

    RegContador = CStr(RSTB.RecordCount)

    If MGrid1.Rows = 2 Then
    Me.Caption = "Buscar Cliente - " & RegContador & " clientes encontrados"
    Else
    Me.Caption = "Buscar Cliente - " & RegContador & " clientes encontrados"
    End If
    Disconnection

End Function


Private Function CarregaFlex1()

MGrid1.Rows = 2

Connection

    RSTB.Open "SELECT * FROM Loc WHERE Nome Like '%" & Text1.Text & "%'", CONE, adOpenStatic, adLockOptimistic

    Do While Not RSTB.EOF
    If RSTB.Fields(0).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 0) = RSTB.Fields(0).Value
    If RSTB.Fields(1).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 1) = RSTB.Fields(1).Value
    If RSTB.Fields(2).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 2) = RSTB.Fields(2).Value
    If RSTB.Fields(3).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 3) = RSTB.Fields(3).Value
    If RSTB.Fields(4).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 4) = RSTB.Fields(4).Value
    If RSTB.Fields(5).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 5) = RSTB.Fields(5).Value
    If RSTB.Fields(6).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 6) = RSTB.Fields(6).Value
    If RSTB.Fields(7).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 7) = RSTB.Fields(7).Value
    If RSTB.Fields(8).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 8) = RSTB.Fields(8).Value
    If RSTB.Fields(9).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 9) = RSTB.Fields(9).Value
    If RSTB.Fields(10).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 10) = RSTB.Fields(10).Value
    If RSTB.Fields(11).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 11) = RSTB.Fields(11).Value
    If RSTB.Fields(12).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 12) = RSTB.Fields(12).Value
    If RSTB.Fields(13).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 13) = RSTB.Fields(13).Value
    If RSTB.Fields(14).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 14) = RSTB.Fields(14).Value
    If RSTB.Fields(15).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 15) = RSTB.Fields(15).Value
    If RSTB.Fields(16).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 16) = RSTB.Fields(16).Value
    If RSTB.Fields(17).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 17) = RSTB.Fields(17).Value
    If RSTB.Fields(18).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 18) = RSTB.Fields(18).Value
    If RSTB.Fields(19).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 19) = RSTB.Fields(19).Value
    If RSTB.Fields(20).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 20) = RSTB.Fields(20).Value
    If RSTB.Fields(21).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 21) = RSTB.Fields(21).Value
    If RSTB.Fields(22).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 22) = RSTB.Fields(22).Value
    If RSTB.Fields(23).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 23) = RSTB.Fields(23).Value
    If RSTB.Fields(24).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 24) = RSTB.Fields(24).Value
       
    MGrid1.Rows = MGrid1.Rows + 1
    RSTB.MoveNext
    
    Loop

    MGrid1.Rows = MGrid1.Rows - 1

    RegContador = CStr(RSTB.RecordCount)

    If MGrid1.Rows = 2 Then
    Me.Caption = "Buscar Cliente - " & RegContador & " clientes encontrados"
    Else
    Me.Caption = "Buscar Cliente - " & RegContador & " clientes encontrados"
    End If
    Disconnection
End Function


Private Function CarregaFlex2()
MGrid1.Rows = 2

Connection

    RSTB.Open "SELECT * FROM Contrato WHERE Codigo Like '%" & Text1.Text & "%'", CONE, adOpenStatic, adLockOptimistic

    Do While Not RSTB.EOF
    If RSTB.Fields(0).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 0) = RSTB.Fields(0).Value
    If RSTB.Fields(1).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 1) = RSTB.Fields(1).Value
    If RSTB.Fields(2).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 2) = RSTB.Fields(2).Value
    If RSTB.Fields(3).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 3) = RSTB.Fields(3).Value
    If RSTB.Fields(4).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 4) = RSTB.Fields(4).Value
    If RSTB.Fields(5).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 5) = RSTB.Fields(5).Value
    If RSTB.Fields(6).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 6) = RSTB.Fields(6).Value
    If RSTB.Fields(7).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 7) = RSTB.Fields(7).Value
    If RSTB.Fields(8).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 8) = RSTB.Fields(8).Value
    If RSTB.Fields(9).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 9) = RSTB.Fields(9).Value
    If RSTB.Fields(10).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 10) = RSTB.Fields(10).Value
    If RSTB.Fields(11).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 11) = RSTB.Fields(11).Value
    If RSTB.Fields(12).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 12) = RSTB.Fields(12).Value
    If RSTB.Fields(13).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 13) = RSTB.Fields(13).Value
    If RSTB.Fields(14).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 14) = RSTB.Fields(14).Value
    If RSTB.Fields(15).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 15) = RSTB.Fields(15).Value
    If RSTB.Fields(16).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 16) = RSTB.Fields(16).Value
    If RSTB.Fields(17).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 17) = RSTB.Fields(17).Value
    If RSTB.Fields(18).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 18) = RSTB.Fields(18).Value
    If RSTB.Fields(19).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 19) = RSTB.Fields(19).Value
    If RSTB.Fields(20).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 20) = RSTB.Fields(20).Value
    If RSTB.Fields(21).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 21) = RSTB.Fields(21).Value
    If RSTB.Fields(22).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 22) = RSTB.Fields(22).Value
    If RSTB.Fields(23).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 23) = RSTB.Fields(23).Value
    If RSTB.Fields(24).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 24) = RSTB.Fields(24).Value
    If RSTB.Fields(25).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 25) = RSTB.Fields(25).Value
    If RSTB.Fields(26).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 26) = RSTB.Fields(26).Value
    If RSTB.Fields(27).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 27) = RSTB.Fields(27).Value
    If RSTB.Fields(28).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 28) = RSTB.Fields(28).Value
    If RSTB.Fields(29).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 29) = RSTB.Fields(29).Value
    If RSTB.Fields(30).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 30) = RSTB.Fields(30).Value
    If RSTB.Fields(31).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 31) = RSTB.Fields(31).Value
    If RSTB.Fields(32).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 32) = RSTB.Fields(32).Value
    If RSTB.Fields(33).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 33) = RSTB.Fields(33).Value
    If RSTB.Fields(34).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 34) = RSTB.Fields(34).Value
    If RSTB.Fields(35).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 35) = RSTB.Fields(35).Value
    If RSTB.Fields(36).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 36) = RSTB.Fields(36).Value
    If RSTB.Fields(37).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 37) = RSTB.Fields(37).Value
    If RSTB.Fields(38).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 38) = RSTB.Fields(38).Value
    If RSTB.Fields(39).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 39) = RSTB.Fields(39).Value
    If RSTB.Fields(40).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 40) = RSTB.Fields(40).Value
    If RSTB.Fields(41).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 41) = RSTB.Fields(41).Value
    If RSTB.Fields(42).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 42) = RSTB.Fields(42).Value
    If RSTB.Fields(43).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 43) = RSTB.Fields(43).Value
    If RSTB.Fields(44).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 44) = RSTB.Fields(44).Value
    If RSTB.Fields(45).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 45) = RSTB.Fields(45).Value
    If RSTB.Fields(46).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 46) = RSTB.Fields(46).Value
    If RSTB.Fields(47).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 47) = RSTB.Fields(47).Value
    If RSTB.Fields(48).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 48) = RSTB.Fields(48).Value
    If RSTB.Fields(49).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 49) = RSTB.Fields(49).Value
    If RSTB.Fields(50).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 50) = RSTB.Fields(50).Value
    If RSTB.Fields(51).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 51) = RSTB.Fields(51).Value
    If RSTB.Fields(52).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 52) = RSTB.Fields(52).Value
    If RSTB.Fields(53).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 53) = RSTB.Fields(53).Value
    If RSTB.Fields(54).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 54) = RSTB.Fields(54).Value
    If RSTB.Fields(55).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 55) = RSTB.Fields(55).Value
    If RSTB.Fields(56).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 56) = RSTB.Fields(56).Value
    If RSTB.Fields(57).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 57) = RSTB.Fields(57).Value
    If RSTB.Fields(58).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 58) = RSTB.Fields(58).Value
    If RSTB.Fields(59).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 59) = RSTB.Fields(59).Value
    If RSTB.Fields(60).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 60) = RSTB.Fields(60).Value
    If RSTB.Fields(61).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 61) = RSTB.Fields(61).Value
    If RSTB.Fields(62).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 62) = RSTB.Fields(62).Value
    If RSTB.Fields(63).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 63) = RSTB.Fields(63).Value
    If RSTB.Fields(64).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 64) = RSTB.Fields(64).Value
    If RSTB.Fields(65).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 65) = RSTB.Fields(65).Value
    If RSTB.Fields(66).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 66) = RSTB.Fields(66).Value
    If RSTB.Fields(67).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 67) = RSTB.Fields(67).Value
    If RSTB.Fields(68).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 68) = RSTB.Fields(68).Value
    If RSTB.Fields(69).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 69) = RSTB.Fields(69).Value
    If RSTB.Fields(70).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 70) = RSTB.Fields(70).Value
    If RSTB.Fields(71).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 71) = RSTB.Fields(71).Value
    If RSTB.Fields(72).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 72) = RSTB.Fields(72).Value
    If RSTB.Fields(73).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 73) = RSTB.Fields(73).Value
    If RSTB.Fields(74).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 74) = RSTB.Fields(74).Value
    If RSTB.Fields(75).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 75) = RSTB.Fields(75).Value
    If RSTB.Fields(76).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 76) = RSTB.Fields(76).Value
    If RSTB.Fields(77).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 77) = RSTB.Fields(77).Value
    If RSTB.Fields(78).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 78) = RSTB.Fields(78).Value
    If RSTB.Fields(79).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 79) = RSTB.Fields(79).Value
    If RSTB.Fields(80).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 80) = RSTB.Fields(80).Value
    If RSTB.Fields(81).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 81) = RSTB.Fields(81).Value
    If RSTB.Fields(82).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 82) = RSTB.Fields(82).Value
    If RSTB.Fields(83).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 83) = RSTB.Fields(83).Value
    If RSTB.Fields(84).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 84) = RSTB.Fields(84).Value
    If RSTB.Fields(85).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 85) = RSTB.Fields(85).Value
    If RSTB.Fields(86).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 86) = RSTB.Fields(86).Value
    If RSTB.Fields(87).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 87) = RSTB.Fields(87).Value
    If RSTB.Fields(88).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 88) = RSTB.Fields(88).Value
    If RSTB.Fields(89).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 89) = RSTB.Fields(89).Value
    If RSTB.Fields(90).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 90) = RSTB.Fields(90).Value
    If RSTB.Fields(91).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 91) = RSTB.Fields(91).Value
    
    MGrid1.Rows = MGrid1.Rows + 1
    RSTB.MoveNext
    
    Loop

    MGrid1.Rows = MGrid1.Rows - 1

    RegContador = CStr(RSTB.RecordCount)

    If MGrid1.Rows = 2 Then
    Me.Caption = "Buscar Cliente - " & RegContador & " clientes encontrados"
    Else
    Me.Caption = "Buscar Cliente - " & RegContador & " clientes encontrados"
    End If
    Disconnection
End Function

Private Function CarregaFlex3()

MGrid1.Rows = 2

Connection

    RSTB.Open "SELECT * FROM Loc WHERE Codigo Like '%" & Text1.Text & "%'", CONE, adOpenStatic, adLockOptimistic

    Do While Not RSTB.EOF
    If RSTB.Fields(0).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 0) = RSTB.Fields(0).Value
    If RSTB.Fields(1).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 1) = RSTB.Fields(1).Value
    If RSTB.Fields(2).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 2) = RSTB.Fields(2).Value
    If RSTB.Fields(3).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 3) = RSTB.Fields(3).Value
    If RSTB.Fields(4).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 4) = RSTB.Fields(4).Value
    If RSTB.Fields(5).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 5) = RSTB.Fields(5).Value
    If RSTB.Fields(6).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 6) = RSTB.Fields(6).Value
    If RSTB.Fields(7).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 7) = RSTB.Fields(7).Value
    If RSTB.Fields(8).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 8) = RSTB.Fields(8).Value
    If RSTB.Fields(9).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 9) = RSTB.Fields(9).Value
    If RSTB.Fields(10).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 10) = RSTB.Fields(10).Value
    If RSTB.Fields(11).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 11) = RSTB.Fields(11).Value
    If RSTB.Fields(12).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 12) = RSTB.Fields(12).Value
    If RSTB.Fields(13).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 13) = RSTB.Fields(13).Value
    If RSTB.Fields(14).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 14) = RSTB.Fields(14).Value
    If RSTB.Fields(15).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 15) = RSTB.Fields(15).Value
    If RSTB.Fields(16).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 16) = RSTB.Fields(16).Value
    If RSTB.Fields(17).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 17) = RSTB.Fields(17).Value
    If RSTB.Fields(18).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 18) = RSTB.Fields(18).Value
    If RSTB.Fields(19).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 19) = RSTB.Fields(19).Value
    If RSTB.Fields(20).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 20) = RSTB.Fields(20).Value
    If RSTB.Fields(21).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 21) = RSTB.Fields(21).Value
    If RSTB.Fields(22).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 22) = RSTB.Fields(22).Value
    If RSTB.Fields(23).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 23) = RSTB.Fields(23).Value
    If RSTB.Fields(24).Value <> "" Then MGrid1.TextMatrix(MGrid1.Rows - 1, 24) = RSTB.Fields(24).Value
       
    MGrid1.Rows = MGrid1.Rows + 1
    RSTB.MoveNext
    
    Loop

    MGrid1.Rows = MGrid1.Rows - 1

    RegContador = CStr(RSTB.RecordCount)

    If MGrid1.Rows = 2 Then
    Me.Caption = "Buscar Cliente - " & RegContador & " clientes encontrados"
    Else
    Me.Caption = "Buscar Cliente - " & RegContador & " clientes encontrados"
    End If
    Disconnection
End Function
