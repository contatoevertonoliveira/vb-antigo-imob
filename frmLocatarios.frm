VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmClientes 
   BorderStyle     =   0  'None
   Caption         =   "Formulário Cadastro - Clientes"
   ClientHeight    =   6825
   ClientLeft      =   510
   ClientTop       =   1845
   ClientWidth     =   10980
   Icon            =   "frmLocatarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmLocatarios.frx":000C
      Left            =   3240
      List            =   "frmLocatarios.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4320
      Width           =   735
   End
   Begin MSMask.MaskEdBox MSKCpf 
      DataField       =   "Cpf"
      DataSource      =   "dtaClientes"
      Height          =   285
      Left            =   4080
      TabIndex        =   13
      Top             =   4320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   14
      Mask            =   "###.###.###-##"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   57
      Top             =   5400
      Width           =   4815
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   58
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Digite algum nome para buscar:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1080
         TabIndex        =   59
         Top             =   240
         Width           =   2685
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmLocatarios.frx":0025
      Height          =   1095
      Left            =   5040
      OleObjectBlob   =   "frmLocatarios.frx":003F
      TabIndex        =   56
      Top             =   5520
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<<  Adicionar  >>"
      Enabled         =   0   'False
      Height          =   735
      Left            =   8760
      Picture         =   "frmLocatarios.frx":0A1A
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtEstCivil 
      Alignment       =   2  'Center
      DataField       =   "EstadoCivil"
      DataSource      =   "dtaClientes"
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Data dtaClientes 
      Caption         =   "Clientes"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programa Imobiliária\Dados\Bdimobiliaria.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clientes"
      Top             =   5040
      Width           =   10695
   End
   Begin VB.TextBox txtCodigo 
      BackColor       =   &H0000FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   53
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Menu:"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   6960
      TabIndex        =   52
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton cmdConfirma 
         Caption         =   "Ir para..."
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
         Left            =   480
         TabIndex        =   27
         Top             =   720
         Width           =   3015
      End
      Begin VB.ComboBox cboMenu 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.TextBox txtRgconjuge 
      Alignment       =   2  'Center
      DataField       =   "RGConjuge"
      DataSource      =   "dtaClientes"
      Height          =   285
      Left            =   8040
      MaxLength       =   14
      TabIndex        =   19
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox txtRg 
      Alignment       =   2  'Center
      DataField       =   "RG"
      DataSource      =   "dtaClientes"
      Height          =   285
      Left            =   840
      MaxLength       =   14
      TabIndex        =   11
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opcões de Controle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   50
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton cmdalterar 
         Caption         =   "&Alterar"
         Enabled         =   0   'False
         Height          =   825
         Left            =   4680
         Picture         =   "frmLocatarios.frx":0E5C
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Alterar Registro"
         Top             =   240
         Width           =   1125
      End
      Begin VB.CommandButton CmdIncluir 
         Caption         =   "&S&alvar"
         Enabled         =   0   'False
         Height          =   825
         Left            =   1080
         Picture         =   "frmLocatarios.frx":129E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Gravar Registro"
         Top             =   240
         Width           =   1005
      End
      Begin VB.CommandButton cmdnovopro 
         Caption         =   "&Novo"
         Enabled         =   0   'False
         Height          =   825
         Left            =   120
         Picture         =   "frmLocatarios.frx":16E0
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Novo Registro"
         Top             =   240
         Width           =   885
      End
      Begin VB.CommandButton CmdExcluir 
         Caption         =   "&Excluir"
         Enabled         =   0   'False
         Height          =   825
         Left            =   3480
         Picture         =   "frmLocatarios.frx":1B22
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Excluir Registro"
         Top             =   240
         Width           =   1125
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Enabled         =   0   'False
         Height          =   825
         Left            =   2160
         Picture         =   "frmLocatarios.frx":1F64
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Cancelar Inclusão"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sa&ir"
         Height          =   825
         Left            =   5880
         Picture         =   "frmLocatarios.frx":23A6
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Sair"
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox txtest 
      Alignment       =   2  'Center
      DataField       =   "Estado"
      DataSource      =   "dtaClientes"
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      MaxLength       =   2
      TabIndex        =   7
      ToolTipText     =   "Estado de Moradia do Cadastrado"
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtemail 
      DataField       =   "Email"
      DataSource      =   "dtaClientes"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   720
      MouseIcon       =   "frmLocatarios.frx":27E8
      MousePointer    =   99  'Custom
      TabIndex        =   14
      ToolTipText     =   " Email de Contato de Cadastrado"
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox txtcidadenotifi 
      Alignment       =   2  'Center
      DataField       =   "CidadeNoti"
      DataSource      =   "dtaClientes"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      MaxLength       =   20
      TabIndex        =   6
      Text            =   " "
      ToolTipText     =   "Cidade de Moradia do Cadastrado"
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox txtbairronotifi 
      DataField       =   "BairroNoti"
      DataSource      =   "dtaClientes"
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      MaxLength       =   30
      TabIndex        =   5
      Text            =   " "
      ToolTipText     =   "Nome do Bairro de Localização do Cadastrado"
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox txtruanot 
      DataField       =   "RuaNoti"
      DataSource      =   "dtaClientes"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   4
      Text            =   " "
      ToolTipText     =   "Nome da Rua de Localização de Cadastrado"
      Top             =   2880
      Width           =   5055
   End
   Begin VB.TextBox txtlocatario 
      DataField       =   "Nome"
      DataSource      =   "dtaClientes"
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      MaxLength       =   80
      TabIndex        =   0
      Text            =   " "
      ToolTipText     =   "Nome do Cadastrado - Campo Obrigatorio"
      Top             =   1800
      Width           =   5295
   End
   Begin VB.TextBox txtcod 
      BackColor       =   &H0080C0FF&
      DataField       =   "ID"
      DataSource      =   "dtaClientes"
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      MaxLength       =   5
      TabIndex        =   34
      Text            =   " "
      ToolTipText     =   "Codigo do Cadastro - Campo Obrigatorio"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox TxtProfissao 
      Alignment       =   2  'Center
      DataField       =   "Profissao"
      DataSource      =   "dtaClientes"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4560
      MaxLength       =   50
      TabIndex        =   3
      Text            =   " "
      ToolTipText     =   "Profissão do Cadastrado"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txtnacional 
      Alignment       =   2  'Center
      DataField       =   "Nacionalidade"
      DataSource      =   "dtaClientes"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   1
      Text            =   " "
      ToolTipText     =   "Nacionalidade do Cadastrado"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtcepnotifi 
      Alignment       =   2  'Center
      DataField       =   "CepNoti"
      DataSource      =   "dtaClientes"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4320
      MaxLength       =   9
      TabIndex        =   8
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox txttelefores 
      Alignment       =   2  'Center
      DataField       =   "TelRes"
      DataSource      =   "dtaClientes"
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      MaxLength       =   14
      TabIndex        =   9
      ToolTipText     =   "Numero de Telefone Residencial para contato"
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox txttelefonecom 
      Alignment       =   2  'Center
      DataField       =   "TelCom"
      DataSource      =   "dtaClientes"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4320
      MaxLength       =   14
      TabIndex        =   10
      ToolTipText     =   "Numero de Telefone Comercial para Contato"
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox txtsite 
      DataField       =   "Site"
      DataSource      =   "dtaClientes"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4200
      MouseIcon       =   "frmLocatarios.frx":2AF2
      TabIndex        =   15
      ToolTipText     =   "Endereço Eletronico"
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Frame Frame5 
      Caption         =   "Dados Cadastrais Conjuge"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   6360
      TabIndex        =   28
      Top             =   2520
      Width           =   4455
      Begin MSMask.MaskEdBox mskCpfConj 
         DataField       =   "CPFConjuge"
         DataSource      =   "dtaClientes"
         Height          =   285
         Left            =   1680
         TabIndex        =   60
         Top             =   1920
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   14
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtconjuge 
         DataField       =   "Conjuge"
         DataSource      =   "dtaClientes"
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         MaxLength       =   80
         TabIndex        =   16
         Text            =   " "
         ToolTipText     =   "Nome Conjuge"
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtconjugeNacional 
         Alignment       =   2  'Center
         DataField       =   "NacionalConjuge"
         DataSource      =   "dtaClientes"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   17
         Text            =   " "
         ToolTipText     =   "Nacionalidade Conjuge"
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtProfissaoConjuge 
         Alignment       =   2  'Center
         DataField       =   "ProfConjuge"
         DataSource      =   "dtaClientes"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   18
         Text            =   " "
         ToolTipText     =   "Profissão Conjuge"
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmLocatarios.frx":2DFC
         TabIndex        =   33
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Nacionalidade:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmLocatarios.frx":36C6
         TabIndex        =   32
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Profissão:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmLocatarios.frx":3F90
         TabIndex        =   31
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "CPF:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmLocatarios.frx":485A
         TabIndex        =   30
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "R.G:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmLocatarios.frx":5124
         TabIndex        =   29
         Top             =   1560
         Width           =   375
      End
   End
   Begin VB.Line Line2 
      X1              =   6480
      X2              =   8520
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      X1              =   6360
      X2              =   8640
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Adicionar dados ao contrato"
      Height          =   195
      Left            =   8760
      TabIndex        =   55
      Top             =   1440
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEEBEF&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000011&
      Height          =   195
      Left            =   3480
      TabIndex        =   51
      Top             =   1440
      Width           =   75
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      MouseIcon       =   "frmLocatarios.frx":59EE
      TabIndex        =   49
      ToolTipText     =   "Limpar Campo"
      Top             =   3600
      Width           =   645
   End
   Begin VB.Label lblemail 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      MouseIcon       =   "frmLocatarios.frx":62B8
      TabIndex        =   48
      ToolTipText     =   "Limpar Campo"
      Top             =   4680
      Width           =   540
   End
   Begin VB.Label txttelcom 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Tel. Com.:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      MouseIcon       =   "frmLocatarios.frx":6B82
      TabIndex        =   47
      ToolTipText     =   "Limpar Campo"
      Top             =   3960
      Width           =   915
   End
   Begin VB.Label lbltelres 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Tel. Res.:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      MouseIcon       =   "frmLocatarios.frx":744C
      TabIndex        =   46
      ToolTipText     =   "Limpar Campo"
      Top             =   3960
      Width           =   825
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Cep:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      MouseIcon       =   "frmLocatarios.frx":7D16
      TabIndex        =   45
      ToolTipText     =   "Limpar Campo"
      Top             =   3600
      Width           =   420
   End
   Begin VB.Label lblruanot 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      MouseIcon       =   "frmLocatarios.frx":85E0
      TabIndex        =   44
      ToolTipText     =   "Limpar Campo"
      Top             =   2880
      Width           =   870
   End
   Begin VB.Label lblcidadenoti 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      MouseIcon       =   "frmLocatarios.frx":8EAA
      TabIndex        =   43
      ToolTipText     =   "Limpar Campo"
      Top             =   3240
      Width           =   675
   End
   Begin VB.Label lblbairronot 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      MouseIcon       =   "frmLocatarios.frx":9774
      TabIndex        =   42
      ToolTipText     =   "Limpar Campo"
      Top             =   3240
      Width           =   600
   End
   Begin VB.Label lblcod 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   41
      ToolTipText     =   "Limpar Campo"
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "R.G:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      MouseIcon       =   "frmLocatarios.frx":A03E
      TabIndex        =   40
      ToolTipText     =   "Limpar Campo"
      Top             =   4320
      Width           =   390
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Profissão:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3600
      MouseIcon       =   "frmLocatarios.frx":A908
      TabIndex        =   39
      ToolTipText     =   "Limpar Campo"
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Nacionalidade:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmLocatarios.frx":B1D2
      TabIndex        =   38
      ToolTipText     =   "Limpar Campo"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado Civil:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmLocatarios.frx":BA9C
      TabIndex        =   37
      ToolTipText     =   "Limpar Campo"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblnome 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmLocatarios.frx":C366
      TabIndex        =   36
      ToolTipText     =   "Limpar Campo"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Site:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3600
      MouseIcon       =   "frmLocatarios.frx":CC30
      TabIndex        =   35
      ToolTipText     =   "Limpar Campo"
      Top             =   4680
      Width           =   405
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Banco As DAO.Database
Dim Tabela As DAO.Recordset
Public Sql As String

Private Sub cboMenu_Change()
cboMenu = StrConv(cboMenu, vbUpperCase)
End Sub

Private Sub cboMenu_GotFocus()
cboMenu.BackColor = &HFFFF&
FechaCaixas
End Sub

Private Sub cboMenu_LostFocus()
cboMenu.BackColor = &H80000005
End Sub

Private Sub cmdAlterar_Click()
    
    dtaClientes.Recordset.Edit
    AbreCaixas
    txtlocatario.SetFocus
    cmdAlterar.Enabled = False
    cmdAlterar.BackColor = &H8000000F
    CmdIncluir.Enabled = True
    cmdnovopro.Enabled = False
    cmdExcluir.Enabled = False
    cmdCancelar.Enabled = True

End Sub

Private Sub cmdCancelar_Click()
On Error Resume Next

dtaClientes.Recordset.CancelUpdate
dtaClientes.Refresh
txtCodigo = "KL" & Format(txtcod, "000")
MSKCpf.Text = dtaClientes.Recordset("CPF")

cmdCancelar.Enabled = False
CmdIncluir.Enabled = False
cmdnovopro.Enabled = True
cmdAlterar.Enabled = True
cmdExcluir.Enabled = True
Combo1.ListIndex = 0
FechaCaixas

End Sub

Private Sub cmdConfirma_Click()
If cboMenu.Text = "Contratos" Then
End If
If cboMenu.Text = "Em alerta!" Then
End If
If cboMenu.Text = "Emissão de Recibos" Then
frmRecibos.Show
End If
If cboMenu.Text = "Agenda de Compromissos" Then
End If
If cboMenu.Text = "Calculadora" Then
End If
If cboMenu.Text = "Impressora" Then
End If
If cboMenu.Text = "Calendário" Then
End If
End Sub

Private Sub cmdExcluir_Click()
If MsgBox("Confirma Exclusão do Cliente?  -> " & dtaClientes.Recordset![ID], vbQuestion + vbYesNo, "Excluir Clientes") = vbYes Then
   dtaClientes.Recordset.Delete
   dtaClientes.Refresh
End If
End Sub

Private Sub cmdform_Click()
Dim P_objeto As Object

For Each P_objeto In frmClientes.Controls
If TypeOf P_objeto Is TextBox Then
P_objeto.Text = StrConv(P_objeto, vbUpperCase)
End If
Next P_objeto
End Sub

Private Sub cmdincluir_Click()
  
  If MsgBox("Você confirma a gravação dos Dados?", vbYesNo, "Sair do Cadastro de Clientes") = vbYes Then
    dtaClientes.UpdateRecord
    dtaClientes.Recordset.Bookmark = dtaClientes.Recordset.LastModified
    dtaClientes.Refresh
    CmdIncluir.Enabled = False
    cmdnovopro.Enabled = True
    cmdExcluir.Enabled = True
    cmdAlterar.Enabled = True
    cmdCancelar.Enabled = False
    FechaCaixas
    txtcod.Enabled = False
    MsgBox ("Dados cadastrados com sucesso!")
Else
    txtlocatario.SetFocus
    Exit Sub
End If
End Sub

Private Sub cmdnovopro_Click()
On Error Resume Next
Dim Novo As String

Set Banco = OpenDatabase(App.Path & "\Dados\Bdimobiliaria.MDB")
Set Tabela = Banco.OpenRecordset("Clientes", dbOpenTable)

Tabela.Index = "ID"
Tabela.MoveLast
Novo = Tabela!ID
Tabela.Seek "=", Novo
If Tabela.NoMatch = False Then
    Novo = Novo + 1
End If
Banco.Close
    
    dtaClientes.Recordset.AddNew
    LimpaCaixas
    AbreCaixas
    txtcod = Novo
    cmdnovopro.Enabled = False
    cmdAlterar.Enabled = False
    cmdAlterar.BackColor = &H8000000F
    CmdIncluir.Enabled = True
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    txtlocatario.SetFocus
    
End Sub

Private Sub cmdSair_Click()
If MsgBox("Quer sair do Cadastro de Clientes?", vbYesNo, "Sair do Cadastro") = vbYes Then
    Unload frmClientes
    RedefineFormPrincipal
Else
    Exit Sub
End If
frmFundo.Enabled = True
End Sub

Private Sub Combo1_Click()
If Combo1.Text = "CPF" Then
    MSKCpf.Mask = Empty
    MSKCpf.Text = Empty
    MSKCpf.Mask = "###.###.###-##"
ElseIf Combo1.Text = "CNPJ" Then
    MSKCpf.Mask = Empty
    MSKCpf.Text = Empty
    MSKCpf.Mask = "##.###.###/####-##"
End If
End Sub

Private Sub Command1_Click()

If frmContrLoc.Visible = True Then
    If frmContrLoc.ssPainel.Tab = 0 Then
        Tab0
        Unload Me
    ElseIf frmContrLoc.ssPainel.Tab = 1 Then
        Tab1
        Unload Me
    ElseIf frmContrLoc.ssPainel.Tab = 2 Then
        Tab2
        Unload Me
    ElseIf frmContrLoc.ssPainel.Tab = 3 Then
        tab3
        Unload Me
    ElseIf frmContrLoc.ssPainel.Tab = 4 Then
        Tab4
        Unload Me
    End If
End If

If frmCompraVenda.Visible = True Then
    If frmCompraVenda.ssPainel.Tab = 0 Then
        TabC0
        Unload Me
    ElseIf frmCompraVenda.ssPainel.Tab = 1 Then
        TabC1
        Unload Me
    End If
End If
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo erro

If KeyCode = 40 Then

End If

If KeyCode = vbKeyLeft And ActiveControl.SelStart = 0 Then
        SendKeys "+{tab}"
    ElseIf KeyCode = vbKeyRight And ActiveControl.SelStart = Len(ActiveControl.Text) Then
        SendKeys "{tab}"
    End If
erro:
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then
     SendKeys vbTab
     KeyAscii = 0
  End If
  
 If KeyAscii = 27 Then
 If MsgBox("Quer sair do Cadastro de Clientes?", vbYesNo, "Sair do Cadastro") = vbYes Then
    Unload Me
    Banco.Close
Else
    Exit Sub
End If
End If
End Sub

Private Sub Form_Load()

    dtaClientes.DatabaseName = App.Path & "\Dados\Bdimobiliaria.MDB"
    dtaClientes.RecordSource = "Clientes"
    dtaClientes.Refresh
    FechaCaixas
    cmdnovopro.Enabled = True
    cmdAlterar.Enabled = True
    cmdExcluir.Enabled = True
    Combo1.ListIndex = 0
    
End Sub

Private Function FechaCaixas()
Dim P_objeto As Object

For Each P_objeto In Me.Controls
    If TypeOf P_objeto Is TextBox Then
    P_objeto.Enabled = False
    P_objeto.BackColor = &H8000000B
End If
Next P_objeto

Combo1.Enabled = False
Combo1.BackColor = &H8000000B
MSKCpf.Enabled = False
MSKCpf.BackColor = &H8000000B
mskCpfConj.Enabled = False
mskCpfConj.BackColor = &H8000000B
Text1.Enabled = True
Text1.BackColor = &HC0FFC0

End Function

Private Function AbreCaixas()
Dim P_objeto As Object

For Each P_objeto In Me.Controls
    If TypeOf P_objeto Is TextBox Then
    P_objeto.Enabled = True
    P_objeto.BackColor = &H80000005
End If
Next P_objeto

Combo1.Enabled = True
Combo1.BackColor = &H80000005
MSKCpf.Enabled = True
MSKCpf.BackColor = &H80000005
mskCpfConj.Enabled = True
mskCpfConj.BackColor = &H80000005
txtCodigo.Enabled = False
txtcod.Enabled = False

End Function

Private Function LimpaCaixas()

txtcod = ""
txtCodigo = ""
txtlocatario = ""
txtnacional = ""
TxtProfissao = ""
txtruanot = ""
txtbairronotifi = ""
txtcidadenotifi = ""
txtcepnotifi = ""
txtest = ""
txttelefores = ""
txttelefonecom = ""
txtRg = ""
MSKCpf.Mask = Empty
MSKCpf.Text = Empty
MSKCpf.Mask = "###.###.###-##"
mskCpfConj.Mask = Empty
mskCpfConj.Text = Empty
mskCpfConj.Mask = "###.###.###-##"
txtemail = ""
txtsite = ""
txtconjuge = ""
txtconjugeNacional = ""
txtProfissaoConjuge = ""
txtRgconjuge = ""
txtProfissaoConjuge = ""

End Function

Private Function EncheCombo()

With cboMenu
            .AddItem "Contratos"
            .AddItem "Em alerta!"
            .AddItem "Consultas & Pesquisas"
            .AddItem "Emissão de Recibos"
            .AddItem "---------------------"
            .AddItem "Agenda de Compromissos"
            .AddItem "Calculadora"
            .AddItem "Impressora"
            .AddItem "Calendário"
End With

End Function

Private Sub Form_Unload(Cancel As Integer)
frmFundo.Enabled = True
End Sub

Private Sub MSKCpf_Change()
If Len(MSKCpf) = 14 Then
    Combo1.ListIndex = 0
ElseIf Len(MSKCpf) = 18 Then
    Combo1.ListIndex = 1
End If
End Sub

Private Sub mskcpf_LostFocus()
If Combo1.Text = "CPF" Then
    If MSKCpf.Text <> "___.___.___-__" Then
        If VerifCPF(MSKCpf.Text) = False Then
            MsgBox "CPF incorreto.", vbExclamation
            MSKCpf.SelStart = 0
            MSKCpf.SelLength = Len(MSKCpf.Text)
            MSKCpf.SetFocus
            Exit Sub
        End If
    End If
ElseIf Combo1.Text = "CNPJ" Then
    If MSKCpf.Text <> "__.___.___/____-__" Then
    
        Dim Part1 As String
        Dim Part2 As String
        Dim Part3 As String
        Dim Part4 As String
        Dim Part5 As String
        Dim CNPJ As String
        
        Part1 = Mid(Trim(MSKCpf.Text), 1, 2)
        Part2 = Mid(Trim(MSKCpf.Text), 4, 3)
        Part3 = Mid(Trim(MSKCpf.Text), 8, 3)
        Part4 = Mid(Trim(MSKCpf.Text), 12, 4)
        Part5 = Mid(Trim(MSKCpf.Text), 17, 2)
        
        CNPJ = Part1 & Part2 & Part3 & Part4 & Part5
        
        If CNPJValido(CNPJ) = False Then
            MsgBox "CNPJ incorreto.", vbExclamation
            MSKCpf.SelStart = 0
            MSKCpf.SelLength = Len(MSKCpf.Text)
            MSKCpf.SetFocus
            Exit Sub
        End If
        
    End If
End If
End Sub

Private Sub mskCpfConj_LostFocus()
If mskCpfConj.Text <> "___.___.___-__" Then
        If VerifCPF(mskCpfConj.Text) = False Then
            MsgBox "CPF incorreto.", vbExclamation
            mskCpfConj.SelStart = 0
            mskCpfConj.SelLength = Len(mskCpfConj.Text)
            mskCpfConj.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub Text1_Change()
On Error Resume Next
If Text1.Text = "" Then
    dtaClientes.RecordSource = "SELECT * FROM CLIENTES"
    dtaClientes.Refresh
    Exit Sub
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If Text1.Text = "" Then
    dtaClientes.RecordSource = "SELECT * FROM CLIENTES"
    dtaClientes.Refresh
    Exit Sub
End If

dtaClientes.RecordSource = "SELECT * FROM CLIENTES WHERE NOME Like '" & Text1.Text & "*'"
dtaClientes.Refresh
End Sub

Private Sub txtbairronotifi_GotFocus()
txtbairronotifi.BackColor = &HFFFF&
End Sub

Private Sub txtbairronotifi_LostFocus()
txtbairronotifi = StrConv(txtbairronotifi, vbUpperCase)
txtbairronotifi.BackColor = &H80000005
If txtbairronotifi = "" Then
    txtbairronotifi = " "
End If
End Sub

Private Sub txtcepnotifi_GotFocus()
txtcepnotifi.BackColor = &HFFFF&
End Sub

Private Sub txtcepnotifi_LostFocus()
txtcepnotifi.BackColor = &H80000005
txtcepnotifi = StrConv(txtcepnotifi, vbUpperCase)
If txtcepnotifi = "" Then
    txtcepnotifi = " "
End If
End Sub

Private Sub txtcidadenotifi_GotFocus()
txtcidadenotifi.BackColor = &HFFFF&
End Sub

Private Sub txtcidadenotifi_LostFocus()
txtcidadenotifi.BackColor = &H80000005
txtcidadenotifi = StrConv(txtcidadenotifi, vbUpperCase)
If txtcidadenotifi = "" Then
    txtcidadenotifi = " "
End If
End Sub

Private Sub txtcod_Change()
txtCodigo = "KL" & Format(txtcod, "000")
End Sub

Private Sub txtconjuge_GotFocus()
txtconjuge.BackColor = &HFFFF&
End Sub

Private Sub txtconjuge_LostFocus()
txtconjuge.BackColor = &H80000005
txtconjuge = StrConv(txtconjuge, vbUpperCase)
If txtconjuge = "" Then
    txtconjuge = " "
End If
End Sub

Private Sub txtconjugeNacional_GotFocus()
txtconjugeNacional.BackColor = &HFFFF&
End Sub

Private Sub txtconjugeNacional_LostFocus()
txtconjugeNacional.BackColor = &H80000005
txtconjugeNacional = StrConv(txtconjugeNacional, vbUpperCase)
If txtconjugeNacional = "" Then
    txtconjugeNacional = " "
End If
End Sub

Private Sub txtemail_GotFocus()
txtemail.BackColor = &HFFFF&
End Sub

Private Sub txtemail_LostFocus()
txtemail.BackColor = &H80000005
txtemail = StrConv(txtemail, vbLowerCase)
If txtemail = "" Then
    txtemail = " "
End If
End Sub

Private Sub txtest_GotFocus()
txtest.BackColor = &HFFFF&
End Sub

Private Sub txtest_LostFocus()
txtest.BackColor = &H80000005
txtest = StrConv(txtest, vbUpperCase)
If txtest = "" Then
    txtest = " "
End If
End Sub

Private Sub txtEstCivil_gotfocus()
txtEstCivil.BackColor = &HFFFF&
End Sub

Private Sub txtEstCivil_LostFocus()
txtEstCivil.BackColor = &H80000005
txtEstCivil = StrConv(txtEstCivil, vbUpperCase)
If txtEstCivil = "" Then
    txtEstCivil = " "
End If
End Sub

Private Sub txtlocatario_GotFocus()
txtlocatario.BackColor = &HFFFF&
End Sub

Private Sub txtlocatario_LostFocus()
txtlocatario = StrConv(txtlocatario, vbUpperCase)
txtlocatario.BackColor = &H80000005
If txtlocatario = "" Then
    txtlocatario = " "
End If
End Sub

Private Sub txtnacional_GotFocus()
txtnacional.BackColor = &HFFFF&
End Sub

Private Sub txtnacional_LostFocus()
txtnacional = StrConv(txtnacional, vbUpperCase)
txtnacional.BackColor = &H80000005
If txtnacional = "" Then
    txtnacional = " "
End If
End Sub

Private Sub TxtProfissao_GotFocus()
TxtProfissao.BackColor = &HFFFF&
End Sub

Private Sub TxtProfissao_LostFocus()
TxtProfissao = StrConv(TxtProfissao, vbUpperCase)
TxtProfissao.BackColor = &H80000005
If TxtProfissao = "" Then
    TxtProfissao = " "
End If
End Sub

Private Sub txtProfissaoConjuge_GotFocus()
txtProfissaoConjuge.BackColor = &HFFFF&
End Sub

Private Sub txtProfissaoConjuge_LostFocus()
txtProfissaoConjuge.BackColor = &H80000005
txtProfissaoConjuge = StrConv(txtProfissaoConjuge, vbUpperCase)
If txtProfissaoConjuge = "" Then
    txtProfissaoConjuge = " "
End If
End Sub

Private Sub txtRg_GotFocus()
txtRg.BackColor = &HFFFF&
End Sub

Private Sub txtRg_LostFocus()
txtRg.BackColor = &H80000005
If txtRg = "" Then
    txtRg = " "
End If
End Sub

Private Sub txtRgConjuge_GotFocus()
txtRgconjuge.BackColor = &HFFFF&
End Sub

Private Sub txtRgConjuge_LostFocus()
txtRgconjuge.BackColor = &H80000005
If txtRgconjuge = "" Then
    txtRgconjuge = " "
End If
End Sub

Private Sub txtruanot_GotFocus()
txtruanot.BackColor = &HFFFF&
End Sub

Private Sub txtruanot_LostFocus()
txtruanot = StrConv(txtruanot, vbUpperCase)
txtruanot.BackColor = &H80000005
If txtruanot = "" Then
    txtruanot = " "
End If
End Sub

Private Sub txtsite_GotFocus()
txtsite.BackColor = &H80000005
End Sub

Private Sub txtsite_LostFocus()
txtsite.BackColor = &H80000005
txtsite = StrConv(txtsite, vbLowerCase)
If txtsite = "" Then
    txtsite = " "
End If
End Sub

Private Sub txttelefonecom_GotFocus()
txttelefonecom.BackColor = &HFFFF&
End Sub

Private Sub txttelefonecom_LostFocus()
txttelefonecom.BackColor = &H80000005
txttelefonecom = StrConv(txttelefonecom, vbUpperCase)
If txttelefonecom = "" Then
    txttelefonecom = " "
End If
End Sub

Private Sub txttelefores_GotFocus()
txttelefores.BackColor = &HFFFF&
End Sub

Private Sub txttelefores_LostFocus()
txttelefores.BackColor = &H80000005
txttelefores = StrConv(txttelefores, vbUpperCase)
If txttelefores = "" Then
    txttelefores = " "
End If
End Sub

Private Function Tab0()

frmContrLoc.Text3 = txtlocatario
frmContrLoc.Text4 = txtnacional
frmContrLoc.Text5 = TxtProfissao
frmContrLoc.Text89 = txtEstCivil
frmContrLoc.Text6 = txtRg
frmContrLoc.Text7 = MSKCpf
frmContrLoc.Text8 = txtruanot
frmContrLoc.Text9 = txtbairronotifi
frmContrLoc.Text10 = txtcidadenotifi
frmContrLoc.Text11 = txtest
frmContrLoc.Text12 = txtcepnotifi
frmContrLoc.Text13 = txtconjuge
frmContrLoc.Text14 = txtconjugeNacional
frmContrLoc.Text15 = txtProfissaoConjuge
frmContrLoc.Text16 = txtRgconjuge
frmContrLoc.Text17 = mskCpfConj

End Function


Private Function Tab1()

frmContrLoc.Text18 = txtlocatario
frmContrLoc.Text19 = txtnacional
frmContrLoc.Text20 = TxtProfissao
frmContrLoc.Text90 = txtEstCivil
frmContrLoc.Text21 = txtRg
frmContrLoc.Text22 = MSKCpf
frmContrLoc.Text23 = txtruanot
frmContrLoc.Text24 = txtbairronotifi
frmContrLoc.Text25 = txtcidadenotifi
frmContrLoc.Text26 = txtest
frmContrLoc.Text27 = txtcepnotifi
frmContrLoc.Text28 = txtconjuge
frmContrLoc.Text29 = txtconjugeNacional
frmContrLoc.Text30 = txtProfissaoConjuge
frmContrLoc.Text31 = txtRgconjuge
frmContrLoc.Text32 = mskCpfConj

End Function


Private Function Tab2()

frmContrLoc.Text33 = txtlocatario
frmContrLoc.Text34 = txtnacional
frmContrLoc.Text35 = TxtProfissao
frmContrLoc.Text91 = txtEstCivil
frmContrLoc.Text36 = txtRg
frmContrLoc.Text37 = MSKCpf
frmContrLoc.Text38 = txtruanot
frmContrLoc.Text39 = txtbairronotifi
frmContrLoc.Text40 = txtcidadenotifi
frmContrLoc.Text41 = txtest
frmContrLoc.Text42 = txtcepnotifi
frmContrLoc.Text43 = txtconjuge
frmContrLoc.Text44 = txtconjugeNacional
frmContrLoc.Text45 = txtProfissaoConjuge
frmContrLoc.Text46 = txtRgconjuge
frmContrLoc.Text47 = mskCpfConj

End Function

Private Function tab3()

frmContrLoc.Text48 = txtlocatario
frmContrLoc.Text49 = txtnacional
frmContrLoc.Text50 = TxtProfissao
frmContrLoc.Text92 = txtEstCivil
frmContrLoc.Text51 = txtRg
frmContrLoc.Text52 = MSKCpf
frmContrLoc.Text53 = txtruanot
frmContrLoc.Text54 = txtbairronotifi
frmContrLoc.Text55 = txtcidadenotifi
frmContrLoc.Text56 = txtest
frmContrLoc.Text57 = txtcepnotifi
frmContrLoc.Text58 = txtconjuge
frmContrLoc.Text59 = txtconjugeNacional
frmContrLoc.Text60 = txtProfissaoConjuge
frmContrLoc.Text61 = txtRgconjuge
frmContrLoc.Text62 = mskCpfConj

End Function

Private Function Tab4()

frmContrLoc.Text63 = txtlocatario
frmContrLoc.Text64 = txtnacional
frmContrLoc.Text65 = TxtProfissao
frmContrLoc.Text93 = txtEstCivil
frmContrLoc.Text66 = txtRg
frmContrLoc.Text67 = MSKCpf
frmContrLoc.Text68 = txtruanot
frmContrLoc.Text69 = txtbairronotifi
frmContrLoc.Text70 = txtcidadenotifi
frmContrLoc.Text71 = txtest
frmContrLoc.Text72 = txtcepnotifi
frmContrLoc.Text73 = txtconjuge
frmContrLoc.Text74 = txtconjugeNacional
frmContrLoc.Text75 = txtProfissaoConjuge
frmContrLoc.Text76 = txtRgconjuge
frmContrLoc.Text77 = mskCpfConj
End Function

Private Function TabC0()

frmCompraVenda.Text(0) = txtlocatario
frmCompraVenda.Text(1) = txtnacional
frmCompraVenda.Text(2) = TxtProfissao
frmCompraVenda.Text(49) = txtEstCivil
frmCompraVenda.Text(3) = txtRg
frmCompraVenda.Text(4) = MSKCpf
frmCompraVenda.Text(5) = txtruanot
frmCompraVenda.Text(6) = txtbairronotifi
frmCompraVenda.Text(7) = txtcidadenotifi
frmCompraVenda.Text(8) = txtest
frmCompraVenda.Text(9) = txtcepnotifi
frmCompraVenda.Text(20) = txtconjuge
frmCompraVenda.Text(21) = txtconjugeNacional
frmCompraVenda.Text(22) = txtProfissaoConjuge
frmCompraVenda.Text(23) = txtRgconjuge
frmCompraVenda.Text(24) = mskCpfConj

End Function


Private Function TabC1()

frmCompraVenda.Text(10) = txtlocatario
frmCompraVenda.Text(11) = txtnacional
frmCompraVenda.Text(12) = TxtProfissao
frmCompraVenda.Text(50) = txtEstCivil
frmCompraVenda.Text(13) = txtRg
frmCompraVenda.Text(14) = MSKCpf
frmCompraVenda.Text(15) = txtruanot
frmCompraVenda.Text(16) = txtbairronotifi
frmCompraVenda.Text(17) = txtcidadenotifi
frmCompraVenda.Text(18) = txtest
frmCompraVenda.Text(19) = txtcepnotifi
frmCompraVenda.Text(25) = txtconjuge
frmCompraVenda.Text(26) = txtconjugeNacional
frmCompraVenda.Text(27) = txtProfissaoConjuge
frmCompraVenda.Text(28) = txtRgconjuge
frmCompraVenda.Text(29) = mskCpfConj

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
