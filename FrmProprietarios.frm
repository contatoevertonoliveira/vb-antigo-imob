VERSION 5.00
Begin VB.Form FrmProprietarios 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   Caption         =   "Formulário Cadastro - Proprietários"
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10980
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmProprietarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   2040
      TabIndex        =   68
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtRg 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3240
      TabIndex        =   65
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox txtCpf 
      Enabled         =   0   'False
      Height          =   285
      Left            =   600
      TabIndex        =   64
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000018&
      Caption         =   "Menu:"
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
      Left            =   6960
      TabIndex        =   61
      Top             =   120
      Width           =   3855
      Begin VB.ComboBox cboMenu 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   360
         Width           =   3615
      End
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
         TabIndex        =   62
         Top             =   840
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdConsultar 
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      Height          =   735
      Left            =   6960
      Picture         =   "FrmProprietarios.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdP 
      BackColor       =   &H80000009&
      Caption         =   "Pr&óximo"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9480
      Picture         =   "FrmProprietarios.frx":044E
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "Próximo Registro"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdA 
      BackColor       =   &H80000009&
      Caption         =   "&Anterior"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8160
      Picture         =   "FrmProprietarios.frx":0890
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Registro Anterior"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ComboBox CobEstadoCivil 
      DataField       =   "EstadoCivil"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "FrmProprietarios.frx":0CD2
      Left            =   1320
      List            =   "FrmProprietarios.frx":0CD4
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtsite 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   8760
      MouseIcon       =   "FrmProprietarios.frx":0CD6
      TabIndex        =   18
      ToolTipText     =   "Endereço Eletronico"
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox txtbloco 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   11
      Text            =   " "
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox txtnumAP 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      MaxLength       =   3
      TabIndex        =   10
      Text            =   " "
      ToolTipText     =   "Numero do Apartamento"
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox txttelefonecom 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      MaxLength       =   13
      TabIndex        =   16
      ToolTipText     =   "Numero de Telefone Comercial para Contato"
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox txttelefores 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   13
      TabIndex        =   15
      ToolTipText     =   "Numero de Telefone Residencial para contato"
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox txtcepnotifi 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      MaxLength       =   9
      TabIndex        =   14
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox txtregime 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   5
      Text            =   " "
      ToolTipText     =   "Regime"
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtnacional 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   2
      Text            =   " "
      ToolTipText     =   "Nacionalidade do Cadastrado"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox TxtProfissao 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      MaxLength       =   50
      TabIndex        =   3
      Text            =   " "
      ToolTipText     =   "Profissão do Cadastrado"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtcod 
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      MaxLength       =   5
      TabIndex        =   0
      Text            =   " "
      ToolTipText     =   "Codigo do Cadastro - Campo Obrigatorio"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtproprietario 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      MaxLength       =   80
      TabIndex        =   1
      Text            =   " "
      ToolTipText     =   "Nome do Cadastrado - Campo Obrigatorio"
      Top             =   1800
      Width           =   5295
   End
   Begin VB.TextBox txtruanot 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   6
      Text            =   " "
      ToolTipText     =   "Nome da Rua de Localização de Cadastrado"
      Top             =   2880
      Width           =   4575
   End
   Begin VB.TextBox txtnnot 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4800
      MaxLength       =   5
      TabIndex        =   9
      Text            =   " "
      ToolTipText     =   "Numero Notificado do Imovel "
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtcomplenot 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Text            =   " "
      ToolTipText     =   "Nome do Edificio de Moradia do Cadastrado"
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox txtbairronotifi 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      MaxLength       =   30
      TabIndex        =   7
      Text            =   " "
      ToolTipText     =   "Nome do Bairro de Localização do Cadastrado"
      Top             =   3240
      Width           =   4335
   End
   Begin VB.TextBox txtcidadenotifi 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      MaxLength       =   20
      TabIndex        =   12
      Text            =   " "
      ToolTipText     =   "Cidade de Moradia do Cadastrado"
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox txtemail 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6000
      MouseIcon       =   "FrmProprietarios.frx":0FE0
      MousePointer    =   99  'Custom
      TabIndex        =   17
      ToolTipText     =   " Email de Contato de Cadastrado"
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
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
      TabIndex        =   24
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H80000009&
         Caption         =   "Sa&ir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   5880
         Picture         =   "FrmProprietarios.frx":12EA
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Sair"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdContrato 
         BackColor       =   &H80000009&
         Caption         =   "&Opções"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   4920
         Picture         =   "FrmProprietarios.frx":172C
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Gerar Contrato"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H80000009&
         Caption         =   "C&ancelar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   1920
         Picture         =   "FrmProprietarios.frx":1B6E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Cancelar Inclusão"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdExcluir 
         BackColor       =   &H80000009&
         Caption         =   "&Excluir"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   3000
         Picture         =   "FrmProprietarios.frx":1FB0
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Excluir Registro"
         Top             =   240
         Width           =   885
      End
      Begin VB.CommandButton cmdnovopro 
         BackColor       =   &H80000009&
         Caption         =   "&Novo"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   120
         Picture         =   "FrmProprietarios.frx":23F2
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Novo Registro"
         Top             =   240
         Width           =   765
      End
      Begin VB.CommandButton CmdIncluir 
         BackColor       =   &H80000009&
         Caption         =   "&S&alvar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   960
         Picture         =   "FrmProprietarios.frx":2834
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Gravar Registro"
         Top             =   240
         Width           =   885
      End
      Begin VB.CommandButton cmdalterar 
         BackColor       =   &H80000009&
         Caption         =   "&Alterar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   3960
         Picture         =   "FrmProprietarios.frx":2C76
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Alterar Registro"
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.TextBox txtest 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5520
      MaxLength       =   2
      TabIndex        =   13
      ToolTipText     =   "Estado de Moradia do Cadastrado"
      Top             =   4320
      Width           =   615
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000018&
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
      TabIndex        =   47
      Top             =   2520
      Width           =   4455
      Begin VB.TextBox txtRgconjuge 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   67
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtCpfconjuge 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   66
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtProfissaoConjuge 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   50
         Text            =   " "
         ToolTipText     =   "Profissão Conjuge"
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtconjugeNacional 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   49
         Text            =   " "
         ToolTipText     =   "Nacionalidade Conjuge"
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtconjuge 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   80
         TabIndex        =   48
         Text            =   " "
         ToolTipText     =   "Nome Conjuge"
         Top             =   480
         Width           =   3495
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
         MouseIcon       =   "FrmProprietarios.frx":30B8
         TabIndex        =   55
         Top             =   1920
         Width           =   375
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
         MouseIcon       =   "FrmProprietarios.frx":3982
         TabIndex        =   54
         Top             =   1560
         Width           =   495
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
         MouseIcon       =   "FrmProprietarios.frx":424C
         TabIndex        =   53
         Top             =   1200
         Width           =   1095
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
         MouseIcon       =   "FrmProprietarios.frx":4B16
         TabIndex        =   52
         Top             =   840
         Width           =   1575
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
         MouseIcon       =   "FrmProprietarios.frx":53E0
         TabIndex        =   51
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Label Label16 
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
      Height          =   255
      Left            =   8280
      MouseIcon       =   "FrmProprietarios.frx":5CAA
      TabIndex        =   46
      ToolTipText     =   "Limpar Campo"
      Top             =   5040
      Width           =   495
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
      MouseIcon       =   "FrmProprietarios.frx":6574
      TabIndex        =   45
      ToolTipText     =   "Limpar Campo"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblBloco 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Blc:"
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
      Left            =   1440
      MouseIcon       =   "FrmProprietarios.frx":6E3E
      TabIndex        =   44
      ToolTipText     =   "Limpar Campo"
      Top             =   3960
      Width           =   330
   End
   Begin VB.Label lblNumAp 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "AP:"
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
      MouseIcon       =   "FrmProprietarios.frx":7708
      TabIndex        =   43
      ToolTipText     =   "Limpar Campo"
      Top             =   3960
      Width           =   300
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Regime:"
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
      Left            =   3000
      MouseIcon       =   "FrmProprietarios.frx":7FD2
      TabIndex        =   42
      ToolTipText     =   "Limpar Campo"
      Top             =   2520
      Width           =   855
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
      MouseIcon       =   "FrmProprietarios.frx":889C
      TabIndex        =   41
      ToolTipText     =   "Limpar Campo"
      Top             =   2520
      Width           =   1455
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
      MouseIcon       =   "FrmProprietarios.frx":9166
      TabIndex        =   40
      ToolTipText     =   "Limpar Campo"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label10 
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
      Left            =   3360
      MouseIcon       =   "FrmProprietarios.frx":9A30
      TabIndex        =   39
      ToolTipText     =   "Limpar Campo"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label9 
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
      Left            =   2760
      MouseIcon       =   "FrmProprietarios.frx":A2FA
      TabIndex        =   38
      ToolTipText     =   "Limpar Campo"
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label8 
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
      MouseIcon       =   "FrmProprietarios.frx":ABC4
      TabIndex        =   37
      ToolTipText     =   "Limpar Campo"
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label lblcomplnot 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Complemento:"
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
      MouseIcon       =   "FrmProprietarios.frx":B48E
      TabIndex        =   36
      ToolTipText     =   "Limpar Campo"
      Top             =   3600
      Width           =   1275
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
      TabIndex        =   35
      ToolTipText     =   "Limpar Campo"
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblbairronot 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro Notificação:"
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
      MouseIcon       =   "FrmProprietarios.frx":BD58
      TabIndex        =   34
      ToolTipText     =   "Limpar Campo"
      Top             =   3240
      Width           =   2175
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
      Left            =   3120
      MouseIcon       =   "FrmProprietarios.frx":C622
      TabIndex        =   33
      ToolTipText     =   "Limpar Campo"
      Top             =   3960
      Width           =   675
   End
   Begin VB.Label lblruanot 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Rua Notificação:"
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
      MouseIcon       =   "FrmProprietarios.frx":CEEC
      TabIndex        =   32
      ToolTipText     =   "Limpar Campo"
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Cep Notificação:"
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
      MouseIcon       =   "FrmProprietarios.frx":D7B6
      TabIndex        =   31
      ToolTipText     =   "Limpar Campo"
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label numeronot 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "N°:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4320
      MouseIcon       =   "FrmProprietarios.frx":E080
      TabIndex        =   30
      ToolTipText     =   "Limpar Campo"
      Top             =   3600
      Width           =   240
   End
   Begin VB.Label lbltelres 
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
      Height          =   255
      Left            =   120
      MouseIcon       =   "FrmProprietarios.frx":E94A
      TabIndex        =   29
      ToolTipText     =   "Limpar Campo"
      Top             =   4680
      Width           =   2415
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
      MouseIcon       =   "FrmProprietarios.frx":F214
      TabIndex        =   28
      ToolTipText     =   "Limpar Campo"
      Top             =   4680
      Width           =   915
   End
   Begin VB.Label lblemail 
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
      Height          =   255
      Index           =   1
      Left            =   5400
      MouseIcon       =   "FrmProprietarios.frx":FADE
      TabIndex        =   27
      ToolTipText     =   "Limpar Campo"
      Top             =   5040
      Width           =   735
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
      Left            =   4680
      MouseIcon       =   "FrmProprietarios.frx":103A8
      TabIndex        =   26
      ToolTipText     =   "Limpar Campo"
      Top             =   4320
      Width           =   645
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
      Left            =   3360
      TabIndex        =   25
      Top             =   1440
      Width           =   75
   End
End
Attribute VB_Name = "FrmProprietarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BD As DAO.Database
Dim tb As DAO.Recordset
Dim wrdApp As Word.Application
Dim wrdAppF As Word.Application
Dim wrdSelection As Word.Selection
Dim Nome_Arq As String
Dim Meuerro$
Dim X As Integer

Sub trata_erro_Word()
'On Error GoTo erro2
Screen.MousePointer = 0
Set wrdApp = Nothing
Set wrdSelection = Nothing

'Me.Caption = " Contrato"
'erro2:
'MsgBox "Ocorreu um erro durante o processamento " & " - Erro numero : " & Err.Number & "   " & Err.Description & "   " & Meuerro
End Sub

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

Private Sub cmdA_Click()
tb.MovePrevious
If tb.BOF = True Then
tb.MoveLast
End If
    Carrega
    txtconjuge.Enabled = False
    txtconjuge.BackColor = &H8000000B
    txtconjugeNacional.Enabled = False
    txtconjugeNacional.BackColor = &H8000000B
    txtCpfconjuge.Enabled = False
    txtCpfconjuge.BackColor = &H8000000B
    txtRgconjuge.Enabled = False
    txtRgconjuge.BackColor = &H8000000B
    txtProfissaoConjuge.Enabled = False
    txtProfissaoConjuge.BackColor = &H8000000B
End Sub

Private Sub cmdAlterar_Click()
    AbreCaixas
    txtproprietario.SetFocus
    cmdAlterar.Enabled = False
    cmdAlterar.BackColor = &H8000000B
    CmdIncluir.Enabled = True
    cmdnovopro.Enabled = False
    cmdExcluir.Enabled = False
    cmdCancelar.Enabled = True
    cmdConsultar.Enabled = False
    cmdp.Enabled = False
    cmda.Enabled = False
If CobEstadoCivil.Text = "Casado" Then
    txtconjuge.Enabled = True
    txtconjuge.BackColor = &H80000005
    txtconjugeNacional.Enabled = True
    txtconjugeNacional.BackColor = &H80000005
    txtCpfconjuge.Enabled = True
    txtCpfconjuge.BackColor = &H80000005
    txtRgconjuge.Enabled = True
    txtRgconjuge.BackColor = &H80000005
    txtProfissaoConjuge.Enabled = True
    txtProfissaoConjuge.BackColor = &H80000005
Else
    txtconjuge.Enabled = False
    txtconjuge.BackColor = &H8000000B
    txtconjugeNacional.Enabled = False
    txtconjugeNacional.BackColor = &H8000000B
    txtCpfconjuge.Enabled = False
    txtCpfconjuge.BackColor = &H8000000B
    txtRgconjuge.Enabled = False
    txtRgconjuge.BackColor = &H8000000B
    txtProfissaoConjuge.Enabled = False
    txtProfissaoConjuge.BackColor = &H8000000B
End If
End Sub

Private Sub cmdCancelar_Click()
If tb.RecordCount = 0 Then
    cmdExcluir.Enabled = False
    cmdAlterar.Enabled = False
    txtcod = Empty
Else
    cmdExcluir.Enabled = True
    cmdAlterar.Enabled = True
    cmda.Enabled = True
    cmdp.Enabled = True
    tb.MoveLast
    Carrega
End If
cmdAlterar.BackColor = &H80000009
cmdCancelar.Enabled = False
CmdIncluir.Enabled = False
cmdnovopro.Enabled = True
cmdConsultar.Enabled = True
FechaCaixas
End Sub

Private Sub cmdConfirma_Click()
If cboMenu.Text = "Contratos" Then
frmOpContrato.Show 1
End If
If cboMenu.Text = "Clientes" Then
Unload Me
frmClientes.Show
End If
If cboMenu.Text = "Em alerta!" Then
End If
If cboMenu.Text = "Emissão de Recibos" Then
frmRecibos.Show 1
End If
If cboMenu.Text = "Agenda de Compromissos" Then
End If
If cboMenu.Text = "Consultas & Pesquisas" Then
frmBusca.Show 1
End If
If cboMenu.Text = "Calculadora" Then
End If
If cboMenu.Text = "Impressora" Then
End If
If cboMenu.Text = "Calendário" Then
End If
End Sub

Private Sub cmdConsultar_Click()
frmBusca.Show 1
End Sub

Private Sub cmdContrato_Click()
frmOpContrato.Show 1
End Sub

Private Sub Substitui_Var1(Header As String, Data As String, oWord As Object)
    On Error Resume Next
    With oWord.Selection.Find
        .ClearFormatting
        .Text = Header
        .Execute Forward:=True
    End With
    Clipboard.Clear
    Clipboard.SetText (Data)
    oWord.Selection.Paste
    Clipboard.Clear
End Sub
Private Sub cmdExcluir_Click()
If MsgBox("Confirma Exclusão do cliente", vbYesNo) = vbYes Then
tb.Delete
cmdA_Click
End If
If tb.RecordCount = 0 Then
cmdExcluir.Enabled = False
cmdAlterar.Enabled = False
LimpaCaixas
FechaCaixas
MsgBox ("Banco de dados está vazio!")
End If
tb.MoveLast
Label1.Caption = "Total de Proprietários Cadastrados: " & Format(tb!codigo, "000")
End Sub

Private Sub cmdincluir_Click()
    Verifica
End Sub

Private Sub cmdnovopro_Click()
    
    AbreCaixas
    LimpaCaixas
    txtcod.Text = Format(tb!codigo + 1, "000")
    txtCodigo = "KP" & Format(txtcod, "000")
    EncheCombo
    cmdnovopro.Enabled = False
    cmdAlterar.Enabled = False
    CmdIncluir.Enabled = True
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmda.Enabled = False
    cmdp.Enabled = False
    cmdConsultar.Enabled = False
    txtproprietario.SetFocus

txtconjuge.Enabled = False
txtconjuge.BackColor = &H8000000B
txtconjugeNacional.Enabled = False
txtconjugeNacional.BackColor = &H8000000B
txtCpfconjuge.Enabled = False
txtCpfconjuge.BackColor = &H8000000B
txtRgconjuge.Enabled = False
txtRgconjuge.BackColor = &H8000000B
txtProfissaoConjuge.Enabled = False
txtProfissaoConjuge.BackColor = &H8000000B


End Sub

Private Sub cmdP_Click()
tb.MoveNext
If tb.EOF = True Then
tb.MovePrevious
End If
    Carrega
    txtconjuge.Enabled = False
    txtconjuge.BackColor = &H8000000B
    txtconjugeNacional.Enabled = False
    txtconjugeNacional.BackColor = &H8000000B
    txtCpfconjuge.Enabled = False
    txtCpfconjuge.BackColor = &H8000000B
    txtRgconjuge.Enabled = False
    txtRgconjuge.BackColor = &H8000000B
    txtProfissaoConjuge.Enabled = False
    txtProfissaoConjuge.BackColor = &H8000000B
End Sub

Private Sub cmdsair_Click()
If MsgBox("Quer sair do Cadastro de Proprietários?", vbYesNo, "Sair do Cadastro") = vbYes Then
    Unload FrmProprietarios
    BD.Close
Else
    Exit Sub
End If
End Sub

Private Sub CobEstadoCivil_Click()
If CobEstadoCivil = "CASADO" Then
    txtconjuge.Enabled = True
    txtconjuge.BackColor = &H80000005
    txtconjugeNacional.Enabled = True
    txtconjugeNacional.BackColor = &H80000005
    txtCpfconjuge.Enabled = True
    txtCpfconjuge.BackColor = &H80000005
    txtRgconjuge.Enabled = True
    txtRgconjuge.BackColor = &H80000005
    txtProfissaoConjuge.Enabled = True
    txtProfissaoConjuge.BackColor = &H80000005
Else
    txtconjuge.Enabled = False
    txtconjuge.BackColor = &H8000000B
    txtconjugeNacional.Enabled = False
    txtconjugeNacional.BackColor = &H8000000B
    txtCpfconjuge.Enabled = False
    txtCpfconjuge.BackColor = &H8000000B
    txtRgconjuge.Enabled = False
    txtRgconjuge.BackColor = &H8000000B
    txtProfissaoConjuge.Enabled = False
    txtProfissaoConjuge.BackColor = &H8000000B
End If
End Sub

Private Sub CobEstadoCivil_LostFocus()
CobEstadoCivil.BackColor = &H80000005
End Sub

Private Sub Form_Activate()
Me.MousePointer = 0
Me.BackColor = &H80000018
Frame1.BackColor = &H80000018
Frame4.BackColor = &H80000018
Frame5.BackColor = &H80000018
End Sub

Private Sub Form_Deactivate()
Me.BackColor = &H80000001
Frame1.BackColor = &H80000001
Frame4.BackColor = &H80000001
Frame5.BackColor = &H80000001
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
 If MsgBox("Quer sair do Cadastro de Proprietários?", vbYesNo, "Sair do Cadastro") = vbYes Then
    Unload FrmProprietarios
    BD.Close
Else
    Exit Sub
End If
End If
End Sub
Private Sub Form_Load()
FechaCaixas
Set BD = OpenDatabase("\\Maq5\c\Meus documentos\Documentos Backup\Programa Imobiliária\Dados\Bdimobiliaria.mdb")
Set tb = BD.OpenRecordset("Prop", dbOpenTable)
tb.Index = "Indcod"
cmdnovopro.Enabled = True
EncheCombo

If tb.EOF = False Then
    Carrega
    cmdAlterar.Enabled = True
    cmdExcluir.Enabled = True
    cmdp.Enabled = True
    cmda.Enabled = True
    cmdContrato.Enabled = True
    cmdConsultar.Enabled = True
End If
tb.MoveLast
Label1.Caption = "Total de Proprietários Cadastrados: " & Format(tb!codigo, "000")
    txtconjuge.Enabled = False
    txtconjuge.BackColor = &H8000000B
    txtconjugeNacional.Enabled = False
    txtconjugeNacional.BackColor = &H8000000B
    txtCpfconjuge.Enabled = False
    txtCpfconjuge.BackColor = &H8000000B
    txtRgconjuge.Enabled = False
    txtRgconjuge.BackColor = &H8000000B
    txtProfissaoConjuge.Enabled = False
    txtProfissaoConjuge.BackColor = &H8000000B
End Sub

Private Function FechaCaixas()
Dim TexObjeto As Object

For Each TexObjeto In Me.Controls
If TypeOf TexObjeto Is TextBox Then
    TexObjeto.Enabled = False
    TexObjeto.BackColor = &H8000000B
End If
Next TexObjeto

txtCpf.Enabled = False
txtCpf.BackColor = &H8000000B
txtRg.Enabled = False
txtRg.BackColor = &H8000000B
txtCpfconjuge.Enabled = False
txtCpfconjuge.BackColor = &H8000000B
txtRgconjuge.Enabled = False
txtRgconjuge.BackColor = &H8000000B
CobEstadoCivil.Enabled = False
CobEstadoCivil.BackColor = &H8000000B

End Function

Private Function AbreCaixas()
Dim TextObjeto As Object

For Each TextObjeto In Me.Controls
If TypeOf TextObjeto Is TextBox Then
    TextObjeto.Enabled = True
    TextObjeto.BackColor = &H80000005
End If
Next TextObjeto
txtCodigo.Enabled = False
txtcod.Enabled = False
txtCpf.Enabled = True
txtCpf.BackColor = &H80000005
txtRg.Enabled = True
txtRg.BackColor = &H80000005
txtCpfconjuge.Enabled = True
txtCpfconjuge.BackColor = &H80000005
txtRgconjuge.Enabled = True
txtRgconjuge.BackColor = &H80000005
CobEstadoCivil.Enabled = True
CobEstadoCivil.BackColor = &H80000005

End Function

Private Function LimpaCaixas()

txtcod = Empty
txtproprietario = Empty
txtnacional = Empty
TxtProfissao = Empty
txtregime = Empty
txtruanot = Empty
txtbairronotifi = Empty
txtcomplenot = Empty
txtnnot = Empty
txtnumAP = Empty
txtbloco = Empty
txtcidadenotifi = Empty
txtest = Empty
txtcepnotifi = Empty
txttelefores = Empty
txttelefonecom = Empty
txtemail = Empty
txtsite = Empty
txtconjuge = Empty
txtconjugeNacional = Empty
txtProfissaoConjuge = Empty

txtCpfconjuge = ""
txtRgconjuge = ""
txtRg = ""
txtCpf = ""

CobEstadoCivil.Clear
CobEstadoCivil.Clear

End Function

Private Function EncheCombo()

With CobEstadoCivil
                    .AddItem "CASADO"
                    .AddItem "SOLTEIRO"
                    .AddItem "DIVORCIADO"
                    .AddItem "DESQUITADO"
                    .AddItem "VIÚVO"
End With

With cboMenu
            .AddItem "Contratos"
            .AddItem "Clientes"
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

Private Function grava()

tb.AddNew
If txtcod <> "" Then tb("codigo") = txtcod
If txtproprietario <> "" Then tb("nome") = txtproprietario
If txtconjuge <> "" Then tb("conjuge") = txtconjuge
If txtnacional <> "" Then tb("nacionalidade") = txtnacional
If txtconjugeNacional <> "" Then tb("nacionalconjuge") = txtconjugeNacional
If TxtProfissao <> "" Then tb("profissao") = TxtProfissao
If txtProfissaoConjuge <> "" Then tb("profconjuge") = txtProfissaoConjuge
If CobEstadoCivil.ListIndex <> -1 Then tb("estadocivil") = CobEstadoCivil
If txtregime <> "" Then tb("regime") = txtregime
If txtruanot <> "" Then tb("ruanoti") = txtruanot
If txtnnot <> "" Then tb("numeronoti") = txtnnot
If txtcomplenot <> "" Then tb("complemento") = txtcomplenot
If txtest <> "" Then tb("estado") = txtest
If txtbairronotifi <> "" Then tb("bairronoti") = txtbairronotifi
If txtcidadenotifi <> "" Then tb("cidadenoti") = txtcidadenotifi
If txtcepnotifi <> "" Then tb("cepnoti") = txtcepnotifi
If txttelefores <> "" Then tb("telres") = txttelefores
If txtRg <> "" Then tb("rg") = txtRg.Text
If txtRgconjuge <> "" Then tb("rgconjuge") = txtRgconjuge.Text
If txtCpf <> "" Then tb("cpf") = txtCpf.Text
If txtCpfconjuge <> "" Then tb("cpfconjuge") = txtCpfconjuge.Text
If txttelefonecom <> "" Then tb("telcom") = txttelefonecom
If txtemail <> "" Then tb("email") = txtemail
If txtnumAP <> "" Then tb("numap") = txtnumAP
If txtbloco <> "" Then tb("bloco") = txtbloco
If txtsite <> "" Then tb("site") = txtsite
tb.Update

End Function

Private Function Altera()

tb.Edit
If txtproprietario <> "" Then tb("nome") = txtproprietario
If txtconjuge <> "" Then tb("conjuge") = txtconjuge
If txtnacional <> "" Then tb("nacionalidade") = txtnacional
If txtconjugeNacional <> "" Then tb("nacionalconjuge") = txtconjugeNacional
If TxtProfissao <> "" Then tb("profissao") = TxtProfissao
If txtProfissaoConjuge <> "" Then tb("profconjuge") = txtProfissaoConjuge
If CobEstadoCivil.ListIndex <> -1 Then tb("estadocivil") = CobEstadoCivil
If txtregime <> "" Then tb("regime") = txtregime
If txtruanot <> "" Then tb("ruanoti") = txtruanot
If txtnnot <> "" Then tb("numeronoti") = txtnnot
If txtcomplenot <> "" Then tb("complemento") = txtcomplenot
If txtest <> "" Then tb("estado") = txtest
If txtbairronotifi <> "" Then tb("bairronoti") = txtbairronotifi
If txtcidadenotifi <> "" Then tb("cidadenoti") = txtcidadenotifi
If txtcepnotifi <> "" Then tb("cepnoti") = txtcepnotifi
If txttelefores <> "" Then tb("telres") = txttelefores
If txtRg <> "" Then tb("rg") = txtRg.Text
If txtRgconjuge <> "" Then tb("rgconjuge") = txtRgconjuge.Text
If txtCpf <> "" Then tb("cpf") = txtCpf.Text
If txtCpfconjuge <> "" Then tb("cpfconjuge") = txtCpfconjuge.Text
If txttelefonecom <> "" Then tb("telcom") = txttelefonecom
If txtemail <> "" Then tb("email") = txtemail
If txtnumAP <> "" Then tb("numap") = txtnumAP
If txtbloco <> "" Then tb("bloco") = txtbloco
If txtsite <> "" Then tb("site") = txtsite
tb.Update

End Function

Private Function Carrega()

If tb("codigo") <> "" Then txtCodigo = tb("codigo")
If tb("codigo") <> "" Then txtcod = tb("codigo")
If tb("nome") <> "" Then txtproprietario = tb("nome")
If tb("conjuge") <> "" Then txtconjuge = tb("conjuge")
If tb("nacionalidade") <> "" Then txtnacional = tb("nacionalidade")
If tb("nacionalconjuge") <> "" Then txtconjugeNacional = tb("nacionalconjuge")
If tb("profissao") <> "" Then TxtProfissao = tb("profissao")
If tb("profconjuge") <> "" Then txtProfissaoConjuge = tb("profconjuge")
If tb("estadocivil") <> "" Then CobEstadoCivil = tb("estadocivil")
If tb("regime") <> "" Then txtregime = tb("regime")
If tb("ruanoti") <> "" Then txtruanot = tb("ruanoti")
If tb("numeronoti") <> "" Then txtnnot = tb("numeronoti")
If tb("complemento") <> "" Then txtcomplenot = tb("complemento")
If tb("estado") <> "" Then txtest = tb("estado")
If tb("bairronoti") <> "" Then txtbairronotifi = tb("bairronoti")
If tb("cidadenoti") <> "" Then txtcidadenotifi = tb("cidadenoti")
If tb("cepnoti") <> "" Then txtcepnotifi = tb("cepnoti")
If tb("telres") <> "" Then txttelefores = tb("telres")
If tb("rg") <> "" Then txtRg = tb("rg")
If tb("rgconjuge") <> "" Then txtRgconjuge = tb("rgconjuge")
If tb("cpf") <> "" Then txtCpf = tb("cpf")
If tb("cpfconjuge") <> "" Then txtCpfconjuge = tb("cpfconjuge")
If tb("telcom") <> "" Then txttelefonecom = tb("telcom")
If tb("email") <> "" Then txtemail = tb("email")
If tb("numap") <> "" Then txtnumAP = tb("numap")
If tb("bloco") <> "" Then txtbloco = tb("bloco")
If tb("site") <> "" Then txtsite = tb("site")

End Function

Private Sub txtbairronotifi_GotFocus()
txtbairronotifi.BackColor = &HFFFF&
End Sub

Private Sub txtbairronotifi_LostFocus()
txtbairronotifi.BackColor = &H80000005
txtbairronotifi = StrConv(txtbairronotifi, vbUpperCase)
End Sub

Private Sub txtbloco_GotFocus()
txtbloco.BackColor = &HFFFF&
End Sub

Private Sub txtbloco_LostFocus()
txtbloco.BackColor = &H80000005
txtbloco = StrConv(txtbloco, vbUpperCase)
End Sub

Private Sub txtcepnotifi_GotFocus()
txtcepnotifi.BackColor = &HFFFF&
End Sub

Private Sub txtcepnotifi_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
        Case Is = 8
        Case 48 To 57
        Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txtcepnotifi_LostFocus()
txtcepnotifi.BackColor = &H80000005
txtcepnotifi = Format(txtcepnotifi, "00000-000")
End Sub

Private Sub txtcidadenotifi_GotFocus()
txtcidadenotifi.BackColor = &HFFFF&
End Sub

Private Sub txtcidadenotifi_LostFocus()
txtcidadenotifi.BackColor = &H80000005
txtcidadenotifi = StrConv(txtcidadenotifi, vbUpperCase)
txtcepnotifi = Format(txtcepnotifi, "00000-000")
End Sub

Private Sub txtcomplenot_GotFocus()
txtcomplenot.BackColor = &HFFFF&
End Sub

Private Sub txtcomplenot_LostFocus()
txtcomplenot.BackColor = &H80000005
txtcomplenot = StrConv(txtcomplenot, vbUpperCase)
End Sub

Private Sub txtconjuge_GotFocus()
txtconjuge.BackColor = &HFFFF&
End Sub

Private Sub txtconjuge_LostFocus()
txtconjuge.BackColor = &H80000005
txtconjuge = StrConv(txtconjuge, vbUpperCase)
End Sub

Private Sub txtconjugeNacional_GotFocus()
txtconjugeNacional.BackColor = &HFFFF&
End Sub

Private Sub txtconjugeNacional_LostFocus()
txtconjugeNacional.BackColor = &H80000005
txtconjugeNacional = StrConv(txtconjugeNacional, vbUpperCase)
End Sub

Private Sub txtcpf_GotFocus()
txtCpf.BackColor = &HFFFF&
End Sub

Private Sub txtcpf_LostFocus()
txtCpf.BackColor = &H80000005
End Sub

Private Sub txtCpfConjuge_GotFocus()
txtCpfconjuge.BackColor = &HFFFF&
End Sub

Private Sub txtCpfConjuge_LostFocus()
txtCpfconjuge.BackColor = &H80000005
End Sub

Private Sub txtemail_GotFocus()
txtemail.BackColor = &HFFFF&
End Sub

Private Sub txtemail_LostFocus()
txtemail.BackColor = &H80000005
txtemail = StrConv(txtemail, vbLowerCase)
End Sub

Private Sub txtest_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
txtest.BackColor = &HFFFF&
End Sub

Private Sub txtest_LostFocus()
txtest.BackColor = &H80000005
txtest = StrConv(txtest, vbUpperCase)
End Sub

Private Sub txtnacional_GotFocus()
txtnacional.BackColor = &HFFFF&
End Sub

Private Sub txtnacional_LostFocus()
txtnacional.BackColor = &H80000005
txtnacional = StrConv(txtnacional, vbUpperCase)
End Sub

Private Sub txtnnot_GotFocus()
txtnnot.BackColor = &HFFFF&
End Sub

Private Sub txtnnot_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
        Case Is = 8
        Case 48 To 57
        Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub txtnnot_LostFocus()
txtnnot.BackColor = &H80000005
End Sub

Private Sub txtnumAP_GotFocus()
txtnumAP.BackColor = &HFFFF&
End Sub

Private Sub txtnumAP_LostFocus()
txtnumAP.BackColor = &H80000005
End Sub

Private Sub TxtProfissao_GotFocus()
TxtProfissao.BackColor = &HFFFF&
End Sub

Private Sub TxtProfissao_LostFocus()
TxtProfissao.BackColor = &H80000005
TxtProfissao = StrConv(TxtProfissao, vbUpperCase)
End Sub

Private Sub txtProfissaoConjuge_GotFocus()
txtProfissaoConjuge.BackColor = &HFFFF&
End Sub

Private Sub txtProfissaoConjuge_LostFocus()
txtProfissaoConjuge.BackColor = &H80000005
txtProfissaoConjuge = StrConv(txtProfissaoConjuge, vbUpperCase)
End Sub

Private Sub txtproprietario_GotFocus()
txtproprietario.BackColor = &HFFFF&
End Sub

Private Sub txtproprietario_LostFocus()
txtproprietario.BackColor = &H80000005
txtproprietario = StrConv(txtproprietario, vbUpperCase)
End Sub

Private Sub txtregime_GotFocus()
txtregime.BackColor = &HFFFF&
End Sub

Private Sub txtregime_LostFocus()
txtregime.BackColor = &H80000005
txtregime = StrConv(txtregime, vbUpperCase)
End Sub

Private Sub txtRg_GotFocus()
txtRg.BackColor = &HFFFF&
End Sub

Private Sub txtRg_LostFocus()
txtRg.BackColor = &H80000005
End Sub

Private Sub txtRgConjuge_GotFocus()
txtRgconjuge.BackColor = &HFFFF&
End Sub

Private Sub txtRgConjuge_LostFocus()
txtRgconjuge.BackColor = &H80000005
End Sub

Private Sub txtruanot_GotFocus()
txtruanot.BackColor = &HFFFF&
End Sub

Private Sub txtruanot_LostFocus()
txtruanot.BackColor = &H80000005
txtruanot = StrConv(txtruanot, vbUpperCase)
End Sub

Private Sub txtsite_GotFocus()
txtsite.BackColor = &HFFFF&
End Sub

Private Sub txtsite_LostFocus()
txtsite.BackColor = &H80000005
txtsite = StrConv(txtsite, vbLowerCase)
End Sub

Private Sub txttelefonecom_GotFocus()
txttelefonecom.BackColor = &HFFFF&
End Sub

Private Sub txttelefonecom_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
        Case Is = 8
        Case 48 To 57
        Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txttelefonecom_LostFocus()
txttelefonecom = Format(txttelefonecom, "(00)0000-0000")
txttelefonecom.BackColor = &H80000005
End Sub

Private Sub txttelefores_GotFocus()
txttelefores.BackColor = &HFFFF&
End Sub

Private Sub txttelefores_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
        Case Is = 8
        Case 48 To 57
        Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txttelefores_LostFocus()
txttelefores = Format(txttelefores, "(00)0000-0000")
txttelefores.BackColor = &H80000005
End Sub

Private Function Verifica()
    
If cmdAlterar.BackColor = &H8000000B Then
Altera
tb.MoveLast
Label1.Caption = "Total de Clientes Cadastrados: " & Format(tb!codigo, "000")
cmda.Enabled = True
cmdp.Enabled = True
CmdIncluir.Enabled = False
cmdnovopro.Enabled = True
cmdExcluir.Enabled = True
cmdAlterar.Enabled = True
cmdCancelar.Enabled = False
FechaCaixas
txtcod.Enabled = False
MsgBox ("Dados cadastrados com sucesso!")
cmdAlterar.BackColor = &H80000009
Else
tb.Index = "Indcod"
tb.Seek "=", txtcod.Text
If tb.NoMatch = False Then
    txtcod = tb!codigo
    MsgBox "Não pode existir usuários com código iguais.", vbCritical
    MsgBox ("Preencha um novo código. Some o atual + 1 e grave!")
    FechaCaixas
    txtcod.Enabled = True
    txtcod.SetFocus
    txtcod.BackColor = &H80C0FF
Else
    grava
    tb.MoveLast
    Label1.Caption = "Total de Clientes Cadastrados: " & Format(tb!codigo, "000")
    cmda.Enabled = True
    cmdp.Enabled = True
    CmdIncluir.Enabled = False
    cmdnovopro.Enabled = True
    cmdExcluir.Enabled = True
    cmdAlterar.Enabled = True
    cmdCancelar.Enabled = False
    FechaCaixas
    txtcod.Enabled = False
    MsgBox ("Dados cadastrados com sucesso!")
End If
End If
End Function
