VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmImoveis 
   BorderStyle     =   0  'None
   Caption         =   "Cadastro de Imóveis"
   ClientHeight    =   6840
   ClientLeft      =   1185
   ClientTop       =   1350
   ClientWidth     =   9600
   Icon            =   "frmImoveis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Buscar Imóvel:"
      Height          =   1695
      Left            =   240
      TabIndex        =   62
      Top             =   4920
      Width           =   9135
      Begin VB.CommandButton Command1 
         Caption         =   "B&uscar Imóvel"
         Height          =   435
         Left            =   120
         TabIndex        =   71
         Top             =   1200
         Width           =   3135
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Bairro"
         Height          =   195
         Left            =   2400
         TabIndex        =   70
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Disponível"
         Height          =   195
         Left            =   1080
         TabIndex        =   69
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Tipo"
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFF80&
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
         TabIndex        =   67
         Top             =   840
         Width           =   3135
      End
      Begin VB.OptionButton Option3 
         Caption         =   "End."
         Height          =   195
         Left            =   2400
         TabIndex        =   66
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Proprietário"
         Height          =   195
         Left            =   1080
         TabIndex        =   65
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Código"
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   360
         Width           =   855
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmImoveis.frx":0442
         Height          =   1335
         Left            =   3360
         OleObjectBlob   =   "frmImoveis.frx":0456
         TabIndex        =   63
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.CommandButton cmdImovel 
      Caption         =   "Adicion&ar Dados"
      Enabled         =   0   'False
      Height          =   420
      Left            =   6000
      TabIndex        =   61
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programa Imobiliária\Dados\Imoveis.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Imoveis"
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      DataField       =   "Cod"
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
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   51
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      DataField       =   "data"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   26
      Left            =   7560
      TabIndex        =   20
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      DataField       =   "fone"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   25
      Left            =   4320
      TabIndex        =   19
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      DataField       =   "prop"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   24
      Left            =   1680
      TabIndex        =   18
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Height          =   420
      Left            =   240
      TabIndex        =   24
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      DataField       =   "captador"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   22
      Left            =   4320
      TabIndex        =   22
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      DataField       =   "valora"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   21
      Left            =   1680
      TabIndex        =   21
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      DataField       =   "valorv"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   23
      Left            =   7560
      TabIndex        =   23
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   420
      Left            =   1440
      TabIndex        =   25
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Enabled         =   0   'False
      Height          =   420
      Left            =   2880
      TabIndex        =   26
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "&Remover"
      Height          =   420
      Left            =   4440
      TabIndex        =   27
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   420
      Left            =   7800
      TabIndex        =   28
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text 
      DataField       =   "ob"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      Index           =   20
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   3120
      Width           =   8895
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      DataField       =   "areacons"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   18
      Left            =   6960
      TabIndex        =   15
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      DataField       =   "est"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   12
      Left            =   8040
      MaxLength       =   2
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      DataField       =   "areau"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   19
      Left            =   8160
      TabIndex        =   16
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      DataField       =   "areat"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   17
      Left            =   5760
      TabIndex        =   14
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      DataField       =   "ladoe"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   16
      Left            =   4560
      TabIndex        =   13
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      DataField       =   "ladod"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   15
      Left            =   3360
      TabIndex        =   12
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      DataField       =   "fund"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   14
      Left            =   2160
      TabIndex        =   11
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      DataField       =   "testp"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   13
      Left            =   360
      TabIndex        =   10
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text 
      DataField       =   "cid"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   11
      Left            =   5040
      TabIndex        =   8
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox Text 
      DataField       =   "bairro"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   2760
      TabIndex        =   7
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text 
      DataField       =   "pontor"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      DataField       =   "cond"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   8040
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      DataField       =   "and"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   7200
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      DataField       =   "elev"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   6360
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      DataField       =   "Bloc"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   5640
      MaxLength       =   5
      TabIndex        =   2
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      DataField       =   "Ap"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   4920
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Text 
      DataField       =   "End"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Frame Frame3 
      Enabled         =   0   'False
      Height          =   3615
      Left            =   240
      TabIndex        =   29
      Top             =   840
      Width           =   9135
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Data da Cap.:"
         Height          =   195
         Left            =   6120
         TabIndex        =   54
         Top             =   2880
         Width           =   990
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Fone:"
         Height          =   195
         Left            =   3240
         TabIndex        =   53
         Top             =   2880
         Width           =   405
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Proprietário:"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   2880
         Width           =   840
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Captador:"
         Height          =   195
         Left            =   3240
         TabIndex        =   50
         Top             =   3240
         Width           =   690
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Valor de Aluguél:"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   3240
         Width           =   1200
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Valor de Venda:"
         Height          =   195
         Left            =   6120
         TabIndex        =   48
         Top             =   3240
         Width           =   1140
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Observação:"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   2040
         Width           =   915
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Área útil:"
         Height          =   195
         Left            =   7920
         TabIndex        =   46
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Área Constr.:"
         Height          =   195
         Left            =   6720
         TabIndex        =   45
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Área Terreno:"
         Height          =   195
         Left            =   5520
         TabIndex        =   44
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Lado esquerdo:"
         Height          =   195
         Left            =   4320
         TabIndex        =   43
         Top             =   1440
         Width           =   1110
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Lado direito:"
         Height          =   195
         Left            =   3120
         TabIndex        =   42
         Top             =   1440
         Width           =   870
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fundos:"
         Height          =   195
         Left            =   1920
         TabIndex        =   41
         Top             =   1440
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Test. Principal:"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   1440
         Width           =   1050
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   7800
         TabIndex        =   39
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         Height          =   195
         Left            =   4800
         TabIndex        =   38
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         Height          =   195
         Left            =   2520
         TabIndex        =   37
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ponto de referência: "
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Condomínio:"
         Height          =   195
         Left            =   7800
         TabIndex        =   35
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Andar:"
         Height          =   195
         Left            =   6960
         TabIndex        =   34
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Elevador:"
         Height          =   195
         Left            =   6120
         TabIndex        =   33
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bloco:"
         Height          =   195
         Left            =   5400
         TabIndex        =   32
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Apto:"
         Height          =   195
         Left            =   4680
         TabIndex        =   31
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   240
      TabIndex        =   55
      Top             =   120
      Width           =   9135
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         DataField       =   "Disponivel"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   2
         Left            =   5760
         TabIndex        =   57
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         DataField       =   "Tipo"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   56
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   720
         TabIndex        =   60
         Top             =   120
         Width           =   540
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Disponibilidade para:"
         Height          =   195
         Left            =   6600
         TabIndex        =   59
         Top             =   120
         Width           =   1470
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Tipo do Imóvel:"
         Height          =   195
         Left            =   3120
         TabIndex        =   58
         Top             =   120
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmImoveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Db As DAO.Database
Dim Tb As DAO.Recordset

Private Sub cmdImprimir_Click()
limpa
End Sub

Private Sub cmdAlterar_Click()

Data1.Recordset.Edit
AbreCaixas
cmdAlterar.Enabled = False
cmdGravar.Enabled = True
cmdNovo.Enabled = False
cmdRemover.Enabled = False
cmdSair.Caption = "&Cancelar"

End Sub

Private Sub cmdGravar_Click()

If Text(1).Text = "" Then
    MsgBox "Informe o Tipo de Imóvel!", vbExclamation, "Gravar Tipo de Imóvel"
    Text(1).SetFocus
    Exit Sub
    End If
If Text(2).Text = "" Then
    MsgBox "Informe a Disposição!", vbExclamation, "Gravar Disposição"
    Text(2).SetFocus
    Exit Sub
    End If
    If Text(3).Text = "" Then
    MsgBox "Informe o Endereço!", vbExclamation, "Gravar Endereço"
    Text(3).SetFocus
    Exit Sub
    End If
    If Text(22).Text = "" Then
    MsgBox "Informe o Captador.", vbExclamation, "Gravar Captador"
    Text(22).SetFocus
    Exit Sub
    End If
    Data1.UpdateRecord
    cmdSair.Caption = "&Sair"
    cmdGravar.Enabled = False
    cmdNovo.Enabled = True
    cmdRemover.Enabled = True
    cmdAlterar.Enabled = True
    MsgBox ("Imóvel cadastrado com sucesso!")
    FechaCaixas
    
End Sub

Private Sub cmdImovel_Click()
If frmContrLoc.Visible = True Then
    frmContrLoc.Text78 = Text(3) & " - " & Text(4) & " - " & Text(5) & " - " & Text(6) & " - " & Text(7)
    frmContrLoc.Text80 = Text(10) & " - " & Text(11) & " - " & Text(12)
    frmContrLoc.Text81 = Text(21)
    Unload Me
End If

If frmCompraVenda.Visible = True Then

    frmCompraVenda.Text(32) = Text(3)
    frmCompraVenda.Text(33) = Text(4)
    frmCompraVenda.Text(34) = Text(5)
    frmCompraVenda.Text(35) = Text(6)
    frmCompraVenda.Text(36) = Text(7)
    frmCompraVenda.Text(37) = Text(8)
    frmCompraVenda.Text(38) = Text(10)
    frmCompraVenda.Text(39) = Text(11)
    frmCompraVenda.Text(40) = Text(12)
    frmCompraVenda.Text(42) = Text(13)
    frmCompraVenda.Text(43) = Text(14)
    frmCompraVenda.Text(44) = Text(15)
    frmCompraVenda.Text(45) = Text(16)
    frmCompraVenda.Text(46) = Text(17)
    frmCompraVenda.Text(47) = Text(18)
    frmCompraVenda.Text(48) = Text(19)
    frmCompraVenda.Text(31) = Text(23)
    Unload Me
    
End If
End Sub

Private Sub cmdNovo_Click()
Dim Novo As String

Set Db = OpenDatabase(App.Path & "\Dados\Imoveis.MDB")
Set Tb = Db.OpenRecordset("Imoveis", dbOpenTable)

Tb.Index = "Cod"
Tb.MoveLast
Novo = Tb!Cod
Tb.Seek "=", Novo
If Tb.NoMatch = False Then
    Novo = Novo + 1
End If
Db.Close
    
    Data1.Recordset.AddNew
    AbreCaixas
    limpa
    Text(0) = Novo
    Text(1).SetFocus
    cmdNovo.Enabled = False
    cmdAlterar.Enabled = False
    cmdRemover.Enabled = False
    cmdGravar.Enabled = True
    cmdSair.Caption = "&Cancelar"
   
End Sub

Private Sub cmdRemover_Click()
If MsgBox("Confirma Exclusão do Imóvel?  -> " & Data1.Recordset![Cod], vbQuestion + vbYesNo, "Excluir Clientes") = vbYes Then
   Data1.Recordset.Delete
   Data1.Refresh
End If
End Sub

Private Sub cmdSair_Click()
If cmdSair.Caption = "&Cancelar" Then
Data1.Recordset.CancelUpdate
Data1.Recordset.MoveLast
cmdNovo.Enabled = True
cmdAlterar.Enabled = True
cmdGravar.Enabled = False
cmdRemover.Enabled = True
FechaCaixas
cmdSair.Caption = "&Sair"
Else
If MsgBox("Quer sair do Cadastro de Imóveis?", vbYesNo, "Sair do sistema") = vbYes Then
    Unload frmImoveis
    RedefineFormPrincipal
  Else
    Exit Sub
End If
End If
frmFundo.Enabled = True
End Sub

Private Sub Command1_Click()
If Option1.Value = True Then
        Text1 = Format(Text1, "000")
    If Text1.Text = "" Then
        Data1.RecordSource = "SELECT * FROM IMOVEIS"
        Data1.Refresh
        Exit Sub
    End If

    Data1.RecordSource = "SELECT * FROM IMOVEIS WHERE COD Like '" & Text1.Text & "*'"
    Data1.Refresh
    
    
ElseIf Option2.Value = True Then
    If Text1.Text = "" Then
        Data1.RecordSource = "SELECT * FROM IMOVEIS"
        Data1.Refresh
        Exit Sub
    End If

    Data1.RecordSource = "SELECT * FROM IMOVEIS WHERE PROP Like '" & Text1.Text & "*'"
    Data1.Refresh
    
    
ElseIf Option3.Value = True Then
    If Text1.Text = "" Then
        Data1.RecordSource = "SELECT * FROM IMOVEIS"
        Data1.Refresh
        Exit Sub
    End If

    Data1.RecordSource = "SELECT * FROM IMOVEIS WHERE END Like '" & Text1.Text & "*'"
    Data1.Refresh
    
    
ElseIf Option4.Value = True Then
    If Text1.Text = "" Then
        Data1.RecordSource = "SELECT * FROM IMOVEIS"
        Data1.Refresh
        Exit Sub
    End If

    Data1.RecordSource = "SELECT * FROM IMOVEIS WHERE TIPO Like '" & Text1.Text & "*'"
    Data1.Refresh
    
    
ElseIf Option5.Value = True Then
    If Text1.Text = "" Then
        Data1.RecordSource = "SELECT * FROM IMOVEIS"
        Data1.Refresh
        Exit Sub
    End If

    Data1.RecordSource = "SELECT * FROM IMOVEIS WHERE DISPONIVEL Like '" & Text1.Text & "*'"
    Data1.Refresh
    
    
ElseIf Option6.Value = True Then
    If Text1.Text = "" Then
        Data1.RecordSource = "SELECT * FROM IMOVEIS"
        Data1.Refresh
        Exit Sub
    End If

    Data1.RecordSource = "SELECT * FROM IMOVEIS WHERE BAIRRO Like '" & Text1.Text & "*'"
    Data1.Refresh
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys ("{TAB}")
    KeyAscii = 0
End If
End Sub

Private Sub Form_Load()

Data1.DatabaseName = App.Path & "\Dados\Imoveis.MDB"
Data1.RecordSource = "Imoveis"

FechaCaixas

End Sub

Private Function FechaCaixas()

Text(1).Enabled = False
Text(1).BackColor = &H8000000F
Text(2).Enabled = False
Text(2).BackColor = &H8000000F
Text(3).Enabled = False
Text(3).BackColor = &H8000000F
Text(4).Enabled = False
Text(4).BackColor = &H8000000F
Text(5).Enabled = False
Text(5).BackColor = &H8000000F
Text(6).Enabled = False
Text(6).BackColor = &H8000000F
Text(7).Enabled = False
Text(7).BackColor = &H8000000F
Text(8).Enabled = False
Text(8).BackColor = &H8000000F
Text(9).Enabled = False
Text(9).BackColor = &H8000000F
Text(10).Enabled = False
Text(10).BackColor = &H8000000F
Text(11).Enabled = False
Text(11).BackColor = &H8000000F
Text(12).Enabled = False
Text(12).BackColor = &H8000000F
Text(13).Enabled = False
Text(13).BackColor = &H8000000F
Text(14).Enabled = False
Text(14).BackColor = &H8000000F
Text(15).Enabled = False
Text(15).BackColor = &H8000000F
Text(16).Enabled = False
Text(16).BackColor = &H8000000F
Text(17).Enabled = False
Text(17).BackColor = &H8000000F
Text(18).Enabled = False
Text(18).BackColor = &H8000000F
Text(19).Enabled = False
Text(19).BackColor = &H8000000F
Text(20).Enabled = False
Text(20).BackColor = &H8000000F
Text(21).Enabled = False
Text(21).BackColor = &H8000000F
Text(22).Enabled = False
Text(22).BackColor = &H8000000F
Text(23).Enabled = False
Text(23).BackColor = &H8000000F
Text(24).Enabled = False
Text(24).BackColor = &H8000000F
Text(25).Enabled = False
Text(25).BackColor = &H8000000F
Text(26).Enabled = False
Text(26).BackColor = &H8000000F

End Function

Private Function AbreCaixas()

Text(1).Enabled = True
Text(1).BackColor = &H80000005
Text(2).Enabled = True
Text(2).BackColor = &H80000005
Text(3).Enabled = True
Text(3).BackColor = &H80000005
Text(4).Enabled = True
Text(4).BackColor = &H80000005
Text(5).Enabled = True
Text(5).BackColor = &H80000005
Text(6).Enabled = True
Text(6).BackColor = &H80000005
Text(7).Enabled = True
Text(7).BackColor = &H80000005
Text(8).Enabled = True
Text(8).BackColor = &H80000005
Text(9).Enabled = True
Text(9).BackColor = &H80000005
Text(10).Enabled = True
Text(10).BackColor = &H80000005
Text(11).Enabled = True
Text(11).BackColor = &H80000005
Text(12).Enabled = True
Text(12).BackColor = &H80000005
Text(13).Enabled = True
Text(13).BackColor = &H80000005
Text(14).Enabled = True
Text(14).BackColor = &H80000005
Text(15).Enabled = True
Text(15).BackColor = &H80000005
Text(16).Enabled = True
Text(16).BackColor = &H80000005
Text(17).Enabled = True
Text(17).BackColor = &H80000005
Text(18).Enabled = True
Text(18).BackColor = &H80000005
Text(19).Enabled = True
Text(19).BackColor = &H80000005
Text(20).Enabled = True
Text(20).BackColor = &H80000005
Text(21).Enabled = True
Text(21).BackColor = &H80000005
Text(22).Enabled = True
Text(22).BackColor = &H80000005
Text(23).Enabled = True
Text(23).BackColor = &H80000005
Text(24).Enabled = True
Text(24).BackColor = &H80000005
Text(25).Enabled = True
Text(25).BackColor = &H80000005
Text(26).Enabled = True
Text(26).BackColor = &H80000005

End Function

Private Sub Form_Unload(Cancel As Integer)
frmFundo.Enabled = True
End Sub

Private Sub Option1_Click()
Text1 = ""
Text1.SetFocus
Data1.Refresh
End Sub

Private Sub Option2_Click()
Text1 = ""
Text1.SetFocus
Data1.Refresh
End Sub

Private Sub Option3_Click()
Text1 = ""
Text1.SetFocus
Data1.Refresh
End Sub

Private Sub Option4_Click()
Text1 = ""
Text1.SetFocus
Data1.Refresh
End Sub

Private Sub Option5_Click()
Text1 = ""
Text1.SetFocus
Data1.Refresh
End Sub

Private Sub Option6_Click()
Text1 = ""
Text1.SetFocus
Data1.Refresh
End Sub

Private Sub Text_Change(Index As Integer)
On Error Resume Next
Text(0) = Format(Text(0), "000")
End Sub

Private Sub Text_LostFocus(Index As Integer)
On Error Resume Next
Select Case Index
Case 1
Text(1) = StrConv(Text(1), vbUpperCase)
Case 2
Text(2) = StrConv(Text(2), vbUpperCase)
Case 3
Text(3) = StrConv(Text(3), vbUpperCase)
Case 4
Text(4) = StrConv(Text(4), vbUpperCase)
Case 5
Text(5) = StrConv(Text(5), vbUpperCase)
Case 6
Text(6) = StrConv(Text(6), vbUpperCase)
Case 7
Text(7) = StrConv(Text(7), vbUpperCase)
Case 8
Text(8) = StrConv(Text(8), vbUpperCase)
Case 9
Text(9) = StrConv(Text(9), vbUpperCase)
Case 10
Text(10) = StrConv(Text(10), vbUpperCase)
Case 11
Text(11) = StrConv(Text(11), vbUpperCase)
Case 12
Text(12) = StrConv(Text(12), vbUpperCase)
Case 13
Text(13) = StrConv(Text(13), vbUpperCase)
Case 14
Text(14) = StrConv(Text(14), vbUpperCase)
Case 15
Text(15) = StrConv(Text(15), vbUpperCase)
Case 16
Text(16) = StrConv(Text(16), vbUpperCase)
Case 17
Text(17) = StrConv(Text(17), vbUpperCase)
Case 18
Text(18) = StrConv(Text(18), vbUpperCase)
Case 19
Text(19) = StrConv(Text(19), vbUpperCase)
Case 20
Text(20) = StrConv(Text(20), vbUpperCase)
Case 21
Text(21) = StrConv(Text(21), vbUpperCase)
Text(21) = Format(Text(21), "Currency")
Case 22
Text(22) = StrConv(Text(22), vbUpperCase)
Case 23
Text(23) = StrConv(Text(23), vbUpperCase)
Text(23) = Format(Text(23), "Currency")
Case 24
Text(24) = StrConv(Text(24), vbUpperCase)
Case 25
Text(25) = StrConv(Text(25), vbUpperCase)
Case 26
Text(26) = CDate(Text(26).Text)
End Select
End Sub

Private Function limpa()

Text(1) = ""
Text(2) = ""
Text(3) = ""
Text(4) = ""
Text(5) = ""
Text(6) = ""
Text(7) = ""
Text(8) = ""
Text(9) = ""
Text(10) = ""
Text(11) = ""
Text(12) = ""
Text(13) = ""
Text(14) = ""
Text(15) = ""
Text(16) = ""
Text(17) = ""
Text(18) = ""
Text(19) = ""
Text(20) = ""
Text(21) = ""
Text(22) = ""
Text(23) = ""
Text(24) = ""
Text(25) = ""
Text(26) = ""

End Function

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
