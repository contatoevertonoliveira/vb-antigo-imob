VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCompraVenda 
   BorderStyle     =   0  'None
   Caption         =   "Super Imob - Compra e Venda"
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame9 
      Caption         =   "Número:"
      Height          =   735
      Left            =   6000
      TabIndex        =   89
      Top             =   600
      Width           =   1095
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   90
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   30
      Left            =   4080
      MaxLength       =   4
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin TabDlg.SSTab ssPainel 
      Height          =   3375
      Left            =   120
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1560
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Vendedor"
      TabPicture(0)   =   "frmCompraVenda.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Comprador"
      TabPicture(1)   =   "frmCompraVenda.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Conjuge"
      TabPicture(2)   =   "frmCompraVenda.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Informações Adicionais"
      TabPicture(3)   =   "frmCompraVenda.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame7 
         Caption         =   "Informações do Imóvel à Venda:"
         Height          =   2655
         Left            =   -74760
         TabIndex        =   84
         Top             =   480
         Width           =   8655
         Begin VB.TextBox Text 
            Height          =   735
            Index           =   32
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   87
            Top             =   1800
            Width           =   8415
         End
         Begin VB.TextBox Text 
            Height          =   735
            Index           =   31
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   85
            Top             =   600
            Width           =   8415
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Informações da Negociação:"
            Height          =   195
            Left            =   120
            TabIndex        =   88
            Top             =   1560
            Width           =   2055
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   120
            TabIndex        =   86
            Top             =   360
            Width           =   765
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Conjuge Comprador:"
         Enabled         =   0   'False
         Height          =   1215
         Left            =   -74760
         TabIndex        =   73
         Top             =   1800
         Width           =   8655
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   25
            Left            =   720
            TabIndex        =   78
            Top             =   360
            Width           =   4575
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   26
            Left            =   6600
            TabIndex        =   77
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   27
            Left            =   960
            TabIndex        =   76
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   28
            Left            =   3720
            TabIndex        =   75
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   29
            Left            =   6480
            TabIndex        =   74
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Left            =   5400
            TabIndex        =   82
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Left            =   120
            TabIndex        =   81
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Rg:"
            Height          =   195
            Left            =   3360
            TabIndex        =   80
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Cpf:"
            Height          =   195
            Left            =   6120
            TabIndex        =   79
            Top             =   720
            Width           =   285
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Conjuge Vendedor:"
         Enabled         =   0   'False
         Height          =   1215
         Left            =   -74760
         TabIndex        =   62
         Top             =   480
         Width           =   8655
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   24
            Left            =   6480
            TabIndex        =   72
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   23
            Left            =   3720
            TabIndex        =   70
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   22
            Left            =   960
            TabIndex        =   67
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   21
            Left            =   6600
            TabIndex        =   65
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   20
            Left            =   720
            TabIndex        =   63
            Top             =   360
            Width           =   4575
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Cpf:"
            Height          =   195
            Left            =   6120
            TabIndex        =   71
            Top             =   720
            Width           =   285
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Rg:"
            Height          =   195
            Left            =   3360
            TabIndex        =   69
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Left            =   120
            TabIndex        =   68
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Left            =   5400
            TabIndex        =   66
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Width           =   465
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados Comprador:"
         Height          =   2295
         Index           =   1
         Left            =   -74760
         TabIndex        =   16
         Top             =   600
         Width           =   8655
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   14
            Left            =   6600
            TabIndex        =   22
            Top             =   1080
            Width           =   1935
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Cnpj:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5880
            TabIndex        =   94
            Top             =   1080
            Width           =   735
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Cpf:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5040
            TabIndex        =   93
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   10
            Left            =   720
            TabIndex        =   17
            Top             =   360
            Width           =   4575
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   11
            Left            =   6600
            TabIndex        =   18
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   12
            Left            =   1080
            TabIndex        =   19
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   6600
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   13
            Left            =   1080
            TabIndex        =   21
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   15
            Left            =   1080
            TabIndex        =   23
            Top             =   1440
            Width           =   4215
         End
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   16
            Left            =   6000
            TabIndex        =   24
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   17
            Left            =   1080
            TabIndex        =   25
            Top             =   1800
            Width           =   2775
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   19
            Left            =   6600
            TabIndex        =   27
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   18
            Left            =   4440
            MaxLength       =   2
            TabIndex        =   26
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Index           =   1
            Left            =   5400
            TabIndex        =   60
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   59
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil:"
            Height          =   195
            Index           =   1
            Left            =   5400
            TabIndex        =   58
            Top             =   720
            Width           =   870
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rg:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   57
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   56
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Index           =   1
            Left            =   5400
            TabIndex        =   55
            Top             =   1440
            Width           =   450
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   54
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Cep:"
            Height          =   195
            Index           =   1
            Left            =   6120
            TabIndex        =   53
            Top             =   1800
            Width           =   330
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Uf:"
            Height          =   195
            Index           =   1
            Left            =   4080
            TabIndex        =   52
            Top             =   1800
            Width           =   210
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados Vendedor:"
         Enabled         =   0   'False
         Height          =   2295
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   8655
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   6600
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1935
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Cnpj:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5880
            TabIndex        =   92
            Top             =   1080
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Cpf:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5040
            TabIndex        =   91
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   360
            Width           =   4575
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   6600
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   6600
            Style           =   2  'Dropdown List
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   5
            Left            =   1080
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1440
            Width           =   4215
         End
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   6
            Left            =   6000
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   7
            Left            =   1080
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1800
            Width           =   2775
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   9
            Left            =   6600
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   8
            Left            =   4440
            MaxLength       =   2
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Index           =   0
            Left            =   5400
            TabIndex        =   50
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   49
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil:"
            Height          =   195
            Index           =   0
            Left            =   5400
            TabIndex        =   48
            Top             =   720
            Width           =   870
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rg:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   47
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Index           =   0
            Left            =   5400
            TabIndex        =   45
            Top             =   1440
            Width           =   450
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   44
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Cep:"
            Height          =   195
            Index           =   0
            Left            =   6120
            TabIndex        =   43
            Top             =   1800
            Width           =   330
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Uf:"
            Height          =   195
            Index           =   0
            Left            =   4080
            TabIndex        =   42
            Top             =   1800
            Width           =   210
         End
      End
   End
   Begin VB.CommandButton Command8 
      Height          =   615
      Left            =   8040
      Picture         =   "frmCompraVenda.frx":0070
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Sair"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Enabled         =   0   'False
      Height          =   615
      Left            =   5400
      Picture         =   "frmCompraVenda.frx":04B2
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Limpar Formulário"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Enabled         =   0   'False
      Height          =   615
      Left            =   4080
      Picture         =   "frmCompraVenda.frx":08F4
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Gravar Contrato"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Enabled         =   0   'False
      Height          =   615
      Left            =   2880
      Picture         =   "frmCompraVenda.frx":0D36
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Excluir Contrato"
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Enabled         =   0   'False
      Height          =   615
      Left            =   1560
      Picture         =   "frmCompraVenda.frx":1178
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Alterar Contrato"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Enabled         =   0   'False
      Height          =   615
      Left            =   240
      Picture         =   "frmCompraVenda.frx":15BA
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Novo Contrato"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      Caption         =   "Código Gerado:"
      Height          =   735
      Left            =   7200
      TabIndex        =   31
      Top             =   600
      Width           =   2055
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Digite o ano do contrato:"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3840
      TabIndex        =   30
      Top             =   600
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Opções:"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   600
      Width           =   3615
      Begin VB.OptionButton Option6 
         Caption         =   "Cadastro de Contrato"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Novo Contrato"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame8 
      Height          =   975
      Left            =   120
      TabIndex        =   39
      Top             =   5040
      Width           =   9135
      Begin VB.CommandButton Command9 
         Enabled         =   0   'False
         Height          =   615
         Left            =   6600
         Picture         =   "frmCompraVenda.frx":19FC
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Buscar Contrato"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label26 
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
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   7320
      TabIndex        =   96
      Top             =   120
      Width           =   885
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   4440
      TabIndex        =   95
      Top             =   120
      Width           =   885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   0
      X2              =   9360
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CONTRATO DE COMPRA E VENDA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   3060
   End
End
Attribute VB_Name = "frmCompraVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dados As DAO.Database
Dim Tabela As DAO.Recordset
Dim Tabelab As DAO.Recordset
Public Sql As String

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 0
Text(3).SetFocus
Case 1
Text(13).SetFocus
End Select
End Sub

Private Sub Command1_Click()
ssPainel.Tab = 1
End Sub

Private Sub Command10_Click()
ssPainel.Tab = 0
End Sub

Private Sub Command2_Click()
ssPainel.Tab = 2
End Sub

Private Sub Command3_Click()
LimpaCaixas
CorTxtAberto
ssPainel.Tab = 0
ssPainel.Enabled = True
Frame1(0).Enabled = True
Frame3.Enabled = True
Frame4.Enabled = True
Option5.Enabled = True
Option6.Enabled = True
Command8.Picture = LoadPicture(App.Path & "\Ícones\TRFFC14.ico")
Command8.ToolTipText = "Cancelar Novo"
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = True
Command7.Enabled = False
Command9.Enabled = False
Text3.BackColor = &H808080
Text1.BackColor = &H808080
End Sub

Private Sub Command4_Click()
ssPainel.Tab = 0
Frame1(0).Enabled = True
AbreTab0
AbreTab1
AbreTab2
AbreTab3
CorTxtAberto
Text(0).SetFocus
Command4.Enabled = False
Command9.Enabled = False
End Sub

Private Sub Command5_Click()
If MsgBox("Confirma Exclusão do cliente", vbYesNo) = vbYes Then
Tabela.Delete
Tabela.MovePrevious
If Tabela.BOF = True Then
Tabela.MoveLast
End If
End If
End Sub

Private Sub Command6_Click()
If Text3 = "" Then
    MsgBox ("Crie o Código!")
    Option6.SetFocus
Else
If Text(30) >= Year(Date) - 3 Then
    GravarA
    Tabela.MoveLast
    MsgBox ("Contrato Novo Cadastrado com Sucesso! O código é " & "KCN" & Format(Tabela!codigo, "000"))
    Tabela.MoveLast
    Label8.Caption = "Contratos Novos: " & Format(Tabela!codigo, "000")
End If

If Text(30) < Year(Date) - 3 Then
    GravarB
    Tabelab.MoveLast
    MsgBox ("Contrato Antigo Cadastrado com Sucesso! O código é " & "KCA" & Format(Tabelab!codigo, "000"))
    Tabelab.MoveLast
    Label26.Caption = "Contratos Antigos: " & Format(Tabelab!codigo, "000")
End If
FechaTab0
FechaTab1
FechaTab2
FechaTab3
CorTxtFechado
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = False
Command7.Enabled = False
Command9.Enabled = True
ssPainel.Tab = 0
Text3.BackColor = &H808080
Text(30).Enabled = False
Command8.ToolTipText = "Sair"
End If
End Sub

Private Sub Command7_Click()
LimpaCaixas
Command7.Enabled = False
Text(30).Enabled = False
Option5.Value = False
Option6.Value = False
End Sub

Private Sub Command8_Click()
If Command8.ToolTipText = "Cancelar Novo" Then
LimpaCaixas
Option5.Value = False
Option6.Value = False
Option5.Enabled = False
Option6.Enabled = False

Command6.Enabled = False
Command8.ToolTipText = "Sair"
Command8.Picture = LoadPicture(App.Path & "\Ícones\ARW10NE.ico")
Command3.Enabled = True
Command7.Enabled = False
FechaTab0
CorTxtFechado
Frame3.Enabled = False
Frame4.Enabled = False
Text1.BackColor = &H808080
If Tabela.EOF = False Then
Command4.Enabled = True
Command5.Enabled = True
Command9.Enabled = True
Tabela.MoveLast
CarregaDadosA
End If
FechaTab0
FechaTab1
FechaTab2
FechaTab3
Command7.Enabled = False
ssPainel.Tab = 0
Text3.Enabled = False
Text3.BackColor = &H808080
Else
If MsgBox("Quer sair do Cadastro de Pessoas?", vbYesNo, "Sair do Cadastro") = vbYes Then
    Unload frmCompraVenda
  Else
    Exit Sub
End If
End If
End Sub

Private Sub Command9_Click()
frmBusca.Show 1
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
End Sub

Private Sub Form_Load()
EncheCombos
Set Dados = OpenDatabase("\\Maq5\c\Meus documentos\Documentos Backup\Programa Imobiliária\Dados\bdimobiliaria.mdb")
Set Tabela = Dados.OpenRecordset("Vcontrato", dbOpenTable)
Set Tabelab = Dados.OpenRecordset("Acontrato", dbOpenTable)
Tabela.Index = "IndCod"
Tabelab.Index = "Indcodigo"

If Tabela.EOF = False Then
CarregaDadosA
Command4.Enabled = True
Command5.Enabled = True
Command9.Enabled = True
End If
Tabela.MoveLast
Label8.Caption = "Contratos Novos: " & Format(Tabela!codigo, "000")
Tabelab.MoveLast
Label26.Caption = "Contratos Antigos: " & Format(Tabelab!codigo, "000")
ssPainel.Tab = 0
Command3.Enabled = True
CorTxtFechado
FechaTab0
Text(30).Enabled = False
Text3.BackColor = &H808080
Text1.BackColor = &H808080
FechaTab1
FechaTab2
FechaTab3
End Sub

Private Sub Option1_Click()
If Len(Text(4)) = 1 Then
    Text(4) = ""
Else
    Text(4).SetFocus
End If
End Sub

Private Sub Option1_KeyPress(KeyAscii As Integer)
Option2.SetFocus
End Sub

Private Sub Option2_Click()
If Len(Text(4)) = 1 Then
    Text(4) = ""
Else
    Text(4).SetFocus
End If
End Sub

Private Sub Option3_Click()
If Len(Text(14)) = 1 Then
    Text(14) = ""
Else
    Text(14).SetFocus
End If
End Sub

Private Sub Option3_KeyPress(KeyAscii As Integer)
Option4.SetFocus
End Sub

Private Sub Option4_Click()
If Len(Text(14)) = 1 Then
    Text(14) = ""
Else
    Text(14).SetFocus
End If
End Sub

Private Sub Option5_Click()
If Option5.Value = True Then
Text(30) = Year(Date)
Text(30).Enabled = True
Text(30).SetFocus
End If
If Option5.Value = True Then
        If Text(30).Text >= Year(Date) Then
            Tabela.Index = "Indcod"
            Tabela.Seek "=", Text3
            If Tabela.NoMatch = False Then
                MsgBox "Não pode existir usuários com código iguais.", vbCritical
                MsgBox ("Preencha um novo código. Some o atual + 1 e grave!")
                Text3.Enabled = True
                Text3.BackColor = &H8000000F
                Text3.ForeColor = &H8000000D
                Text3.SetFocus
            Else
                If Tabela.EOF = True Then
                    Text1 = "001"
                    Text3 = "KCN001"
                    Text(0).SetFocus
                Else
                    Tabela.MoveLast
                    Text1 = Tabela("codigo") + 1
                    Text3 = Format(Text1, "000")
                    Text3 = "KCN" & Format(Text3, "000")
                    Text(0).SetFocus
                End If
            End If
        End If
End If
End Sub

Private Sub Option6_Click()
If Option6.Value = True Then
Text(30).Enabled = True
Text(30).SetFocus
End If
End Sub

Private Sub Text_Change(Index As Integer)
Select Case Index
Case 0
    If Len(Text(0)) = 2 Then
    Command7.Enabled = True
    End If
Case 4
If Len(Text(4)) = 1 Then
    If Option1.Value = False Then
    If Option2.Value = False Then
    Text(4) = ""
    MsgBox ("Selecione uma opção!")
    End If
    End If
End If

If Option1.Value = True Then
    If Len(Text(4)) = 3 Then
    Text(4) = Text(4) + "."
    Text(4).SelStart = 4
    End If

    If Len(Text(4)) = 7 Then
    Text(4) = Text(4) + "."
    Text(4).SelStart = 8
    End If

    If Len(Text(4)) = 11 Then
    Text(4) = Text(4) + "-"
    Text(4).SelStart = 12
    End If
    
    If Len(Text(4)) = 14 Then
    Text(5).SetFocus
    End If
    End If

If Option2.Value = True Then
    If Len(Text(4)) = 2 Then
    Text(4) = Text(4) + "."
    Text(4).SelStart = 3
    End If

    If Len(Text(4)) = 6 Then
    Text(4) = Text(4) + "."
    Text(4).SelStart = 7
    End If

    If Len(Text(4)) = 10 Then
    Text(4) = Text(4) + "/"
    Text(4).SelStart = 11
    End If
    
    If Len(Text(4)) = 15 Then
    Text(4) = Text(4) + "-"
    Text(4).SelStart = 16
    End If
    End If
Case 9
    If Len(Text(9)) = 5 Then
    Text(9) = Text(9) + "-"
    Text(9).SelStart = 6
    End If

Case 14
    If Len(Text(14)) = 1 Then
    If Option3.Value = False Then
    If Option4.Value = False Then
    Text(14) = ""
    MsgBox ("Selecione uma opção!")
    End If
    End If
End If


    If Option3.Value = True Then
    If Len(Text(14)) = 3 Then
    Text(14) = Text(14) + "."
    Text(14).SelStart = 4
    End If

    If Len(Text(14)) = 7 Then
    Text(14) = Text(14) + "."
    Text(14).SelStart = 8
    End If

    If Len(Text(14)) = 11 Then
    Text(14) = Text(14) + "-"
    Text(14).SelStart = 12
    End If
    End If
If Option4.Value = True Then
    If Len(Text(14)) = 2 Then
    Text(14) = Text(14) + "."
    Text(14).SelStart = 3
    End If

    If Len(Text(14)) = 6 Then
    Text(14) = Text(14) + "."
    Text(14).SelStart = 7
    End If

    If Len(Text(14)) = 10 Then
    Text(14) = Text(14) + "/"
    Text(14).SelStart = 11
    End If
    
    If Len(Text(14)) = 15 Then
    Text(14) = Text(14) + "-"
    Text(14).SelStart = 16
    End If
    End If

Case 30
    If Len(Text(30)) = 0 Then
    Text3 = ""
    End If

    If Len(Text(30)) = 4 Then
    If Text(30) < Year(Date) - 3 Then
            Tabelab.Index = "Indcodigo"
            Tabelab.Seek "=", Text3
                If Tabelab.NoMatch = False Then
                    MsgBox "Não pode existir usuários com código iguais.", vbCritical
                    MsgBox ("Preencha um novo código. Some o atual + 1 e grave!")
                    Text3.Enabled = True
                    Text3.BackColor = &H8000000F
                    Text3.ForeColor = &H8000000D
                    Text3.SetFocus
                Else
                    If Tabelab.EOF = True Then
                        Text1 = "001"
                        Text3 = "KCA001"
                        Text(0).SetFocus
                    Else
                        Tabelab.MoveLast
                        Text1 = Format(Tabelab("codigo") + 1, "000")
                        Text3 = "KCA" & Format(Text1, "000")
                        Text(0).SetFocus
                    End If
                End If
        End If
                    
        If Text(30) >= Year(Date) - 3 Then
            Tabela.Index = "Indcod"
            Tabela.Seek "=", Text3
            If Tabela.NoMatch = False Then
                MsgBox "Não pode existir usuários com código iguais.", vbCritical
                MsgBox ("Preencha um novo código. Some o atual + 1 e grave!")
                Text3.Enabled = True
                Text3.BackColor = &H8000000F
                Text3.ForeColor = &H8000000D
                Text3.SetFocus
            Else
                If Tabela.EOF = True Then
                    Text1 = "001"
                    Text3 = "KCN001"
                    Text(0).SetFocus
                Else
                    Tabela.MoveLast
                    Text1 = Tabela("codigo") + 1
                    Text3 = Text1
                    Text3 = "KCN" & Format(Text3, "000")
                    Text(0).SetFocus
                End If
            End If
        End If
    End If
Case 19
    If Len(Text(19)) = 5 Then
    Text(19) = Text(19) + "-"
    Text(19).SelStart = 6
    End If
End Select
End Sub

Private Sub text_GotFocus(Index As Integer)

Select Case Index

Case 0
    Text(0).BackColor = &HFFFF&
Case 1
    Text(1).BackColor = &HFFFF&
Case 2
    Text(2).BackColor = &HFFFF&
Case 3
    Text(3).BackColor = &HFFFF&
Case 4
    Text(4).BackColor = &HFFFF&
Case 5
    Text(5).BackColor = &HFFFF&
Case 6
    Text(6).BackColor = &HFFFF&
Case 7
    Text(7).BackColor = &HFFFF&
Case 8
    Text(8).BackColor = &HFFFF&
Case 9
    Text(9).BackColor = &HFFFF&
Case 10
    Text(10).BackColor = &HFFFF&
Case 11
    Text(11).BackColor = &HFFFF&
Case 12
    Text(12).BackColor = &HFFFF&
Case 13
    Text(13).BackColor = &HFFFF&
Case 14
    Text(14).BackColor = &HFFFF&
Case 15
    Text(15).BackColor = &HFFFF&
Case 16
    Text(16).BackColor = &HFFFF&
Case 17
    Text(17).BackColor = &HFFFF&
Case 18
    Text(18).BackColor = &HFFFF&
Case 19
    Text(19).BackColor = &HFFFF&
Case 20
    Text(20).BackColor = &HFFFF&
Case 21
    Text(21).BackColor = &HFFFF&
Case 22
    Text(22).BackColor = &HFFFF&
Case 23
    Text(23).BackColor = &HFFFF&
Case 24
    Text(24).BackColor = &HFFFF&
Case 25
    Text(25).BackColor = &HFFFF&
Case 26
    Text(26).BackColor = &HFFFF&
Case 27
    Text(27).BackColor = &HFFFF&
Case 28
    Text(28).BackColor = &HFFFF&
Case 29
    Text(29).BackColor = &HFFFF&
Case 30
    Text(30).BackColor = &HFFFF&
Case 31
    Text(31).BackColor = &HFFFF&
Case 32
    Text(32).BackColor = &HFFFF&
End Select
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
Select Case Index
Case 0
    Text(1).SetFocus
Case 1
    Text(2).SetFocus
Case 2
    Combo1(0).SetFocus
Case 3
    Option1.SetFocus
Case 4
    Text(5).SetFocus
Case 5
    Text(6).SetFocus
Case 6
    Text(7).SetFocus
Case 7
    Text(8).SetFocus
Case 8
    Text(9).SetFocus
Case 9
    ssPainel.Tab = 1
    Text(10).SetFocus
Case 10
    Text(11).SetFocus
Case 11
    Text(12).SetFocus
Case 12
    Combo1(1).SetFocus
Case 13
    Option3.SetFocus
Case 14
    Text(15).SetFocus
Case 15
    Text(16).SetFocus
Case 16
    Text(17).SetFocus
Case 17
    Text(18).SetFocus
Case 18
    Text(19).SetFocus
Case 19
    ssPainel.Tab = 2
    Text(20).SetFocus
Case 20
    Text(21).SetFocus
Case 21
    Text(22).SetFocus
Case 22
    Text(23).SetFocus
Case 23
    Text(24).SetFocus
Case 24
    Text(25).SetFocus
Case 25
    Text(26).SetFocus
Case 26
    Text(27).SetFocus
Case 27
    Text(28).SetFocus
Case 28
    Text(29).SetFocus
Case 29
    ssPainel.Tab = 3
    Text(31).SetFocus
Case 31
    Text(32).SetFocus
End Select
End If
End Sub

Private Sub Text_LostFocus(Index As Integer)

Select Case Index

Case 0
    Text(0).BackColor = &H80000005
    Text(0) = StrConv(Text(0), vbUpperCase)
Case 1
    Text(1).BackColor = &H80000005
    Text(1) = StrConv(Text(1), vbUpperCase)
Case 2
    Text(2).BackColor = &H80000005
    Text(2) = StrConv(Text(2), vbUpperCase)
Case 3
    Text(3).BackColor = &H80000005
    Text(3) = StrConv(Text(3), vbUpperCase)
Case 4
    Text(4).BackColor = &H80000005
Case 5
    Text(5).BackColor = &H80000005
    Text(5) = StrConv(Text(5), vbUpperCase)
Case 6
    Text(6).BackColor = &H80000005
    Text(6) = StrConv(Text(6), vbUpperCase)
Case 7
    Text(7).BackColor = &H80000005
    Text(7) = StrConv(Text(7), vbUpperCase)
Case 8
    Text(8).BackColor = &H80000005
    Text(8) = StrConv(Text(8), vbUpperCase)
Case 9
    Text(9).BackColor = &H80000005
Case 10
    Text(10).BackColor = &H80000005
    Text(10) = StrConv(Text(10), vbUpperCase)
Case 11
    Text(11).BackColor = &H80000005
    Text(11) = StrConv(Text(11), vbUpperCase)
Case 12
    Text(12).BackColor = &H80000005
    Text(12) = StrConv(Text(12), vbUpperCase)
Case 13
    Text(13).BackColor = &H80000005
    Text(13) = StrConv(Text(13), vbUpperCase)
Case 14
    Text(14).BackColor = &H80000005
    Text(14) = StrConv(Text(14), vbUpperCase)
Case 15
    Text(15).BackColor = &H80000005
    Text(15) = StrConv(Text(15), vbUpperCase)
Case 16
    Text(16).BackColor = &H80000005
    Text(16) = StrConv(Text(16), vbUpperCase)
Case 17
    Text(17).BackColor = &H80000005
    Text(17) = StrConv(Text(17), vbUpperCase)
Case 18
    Text(18).BackColor = &H80000005
    Text(18) = StrConv(Text(18), vbUpperCase)
Case 19
    Text(19).BackColor = &H80000005
Case 20
    Text(20).BackColor = &H80000005
    Text(20) = StrConv(Text(20), vbUpperCase)
Case 21
    Text(21).BackColor = &H80000005
    Text(21) = StrConv(Text(21), vbUpperCase)
Case 22
    Text(22).BackColor = &H80000005
    Text(22) = StrConv(Text(22), vbUpperCase)
Case 23
    Text(23).BackColor = &H80000005
    Text(23) = StrConv(Text(23), vbUpperCase)
Case 24
    Text(24).BackColor = &H80000005
    Text(24) = StrConv(Text(24), vbUpperCase)
Case 25
    Text(25).BackColor = &H80000005
    Text(25) = StrConv(Text(25), vbUpperCase)
Case 26
    Text(26).BackColor = &H80000005
    Text(26) = StrConv(Text(26), vbUpperCase)
Case 27
    Text(27).BackColor = &H80000005
    Text(27) = StrConv(Text(27), vbUpperCase)
Case 28
    Text(28).BackColor = &H80000005
    Text(28) = StrConv(Text(28), vbUpperCase)
Case 29
    Text(29).BackColor = &H80000005
    Text(29) = StrConv(Text(29), vbUpperCase)
Case 30
    Text(30).BackColor = &H80000005
    If Not Len(Text(30)) = 4 Then
    MsgBox ("Digite um ano válido de 4 dígitos!!")
    Text(30).SetFocus
    End If
Case 31
    Text(31).BackColor = &H80000005
    Text(31) = StrConv(Text(31), vbUpperCase)
Case 32
    Text(32).BackColor = &H80000005
    Text(32) = StrConv(Text(32), vbUpperCase)
End Select
End Sub

Private Function CorTxtFechado()
Dim TxtBox As Object
Dim Combos As Object

For Each TxtBox In Me.Controls
    If TypeOf TxtBox Is TextBox Then
    TxtBox.BackColor = &H8000000B
End If
Next TxtBox

For Each Combos In Me.Controls
    If TypeOf Combos Is ComboBox Then
    Combos.BackColor = &H8000000B
End If
Next Combos

End Function

Private Function CorTxtAberto()
Dim TxtBox As Object
Dim Combos As Object

For Each TxtBox In Me.Controls
    If TypeOf TxtBox Is TextBox Then
    TxtBox.BackColor = &H80000005
End If
Next TxtBox

For Each Combos In Me.Controls
    If TypeOf Combos Is ComboBox Then
    Combos.BackColor = &H80000005
End If
Next Combos

End Function

Private Function AbreTab0()

Text(0).Enabled = True
Text(1).Enabled = True
Text(2).Enabled = True
Text(3).Enabled = True
Text(4).Enabled = True
Text(5).Enabled = True
Text(6).Enabled = True
Text(7).Enabled = True
Text(8).Enabled = True
Text(9).Enabled = True
Combo1(0).Enabled = True
Option1.Enabled = True
Option2.Enabled = True

End Function

Private Function FechaTab0()

Text(0).Enabled = False
Text(1).Enabled = False
Text(2).Enabled = False
Text(3).Enabled = False
Text(4).Enabled = False
Text(5).Enabled = False
Text(6).Enabled = False
Text(7).Enabled = False
Text(8).Enabled = False
Text(9).Enabled = False
Combo1(0).Enabled = False
Option1.Enabled = False
Option1.Value = False
Option2.Enabled = False
Option2.Value = False

End Function

Private Function AbreTab1()

Text(10).Enabled = True
Text(11).Enabled = True
Text(12).Enabled = True
Text(13).Enabled = True
Text(14).Enabled = True
Text(15).Enabled = True
Text(16).Enabled = True
Text(17).Enabled = True
Text(18).Enabled = True
Text(19).Enabled = True
Combo1(1).Enabled = True
Option3.Enabled = True
Option4.Enabled = True

End Function

Private Function FechaTab1()

Text(10).Enabled = False
Text(11).Enabled = False
Text(12).Enabled = False
Text(13).Enabled = False
Text(14).Enabled = False
Text(15).Enabled = False
Text(16).Enabled = False
Text(17).Enabled = False
Text(18).Enabled = False
Text(19).Enabled = False
Combo1(1).Enabled = False
Option3.Enabled = False
Option3.Value = False
Option4.Enabled = False
Option4.Value = False

End Function

Private Function AbreTab2()

Text(20).Enabled = True
Text(21).Enabled = True
Text(22).Enabled = True
Text(23).Enabled = True
Text(24).Enabled = True
Text(25).Enabled = True
Text(26).Enabled = True
Text(27).Enabled = True
Text(28).Enabled = True
Text(29).Enabled = True
Frame2.Enabled = True
Frame6.Enabled = True

End Function

Private Function FechaTab2()

Text(20).Enabled = False
Text(21).Enabled = False
Text(22).Enabled = False
Text(23).Enabled = False
Text(24).Enabled = False
Text(25).Enabled = False
Text(26).Enabled = False
Text(27).Enabled = False
Text(28).Enabled = False
Text(29).Enabled = False
Frame2.Enabled = False
Frame6.Enabled = False

End Function

Private Function AbreTab3()

Text(31).Enabled = True
Text(32).Enabled = True

End Function

Private Function FechaTab3()

Text(31).Enabled = False
Text(32).Enabled = False

End Function

Private Function EncheCombos()

With Combo1(0)
        .AddItem "SOLTEIRO(A)"
        .AddItem "CASADO(A)"
        .AddItem "DIVORCIADO(A)"
        .AddItem "VIÚVO(A)"
End With

With Combo1(1)
        .AddItem "SOLTEIRO(A)"
        .AddItem "CASADO(A)"
        .AddItem "DIVORCIADO(A)"
        .AddItem "VIÚVO(A)"
End With

End Function

Private Function LimpaCaixas()
Dim TextBox As Object

For Each TextBox In Me.Controls
    If TypeOf TextBox Is TextBox Then
    TextBox = Empty
End If
Next TextBox

Combo1(0).ListIndex = -1
Combo1(1).ListIndex = -1

End Function

Private Function GravarA()

Tabela.AddNew

If Text1 <> "" Then Tabela("codigo") = Text1
If Text(0) <> "" Then Tabela("vendedor") = Text(0)
If Text(1) <> "" Then Tabela("nacional") = Text(1)
If Text(2) <> "" Then Tabela("prof") = Text(2)
If Combo1(0) <> "" Then Tabela("estcivil") = Combo1(0)
If Text(3) <> "" Then Tabela("rg") = Text(3)
If Text(4) <> "" Then Tabela("cpf") = Text(4)
If Text(5) <> "" Then Tabela("vend") = Text(5)
If Text(6) <> "" Then Tabela("vbairro") = Text(6)
If Text(7) <> "" Then Tabela("vcidade") = Text(7)
If Text(8) <> "" Then Tabela("vuf") = Text(8)
If Text(9) <> "" Then Tabela("Vcep") = Text(9)
If Text(10) <> "" Then Tabela("comprador") = Text(10)
If Text(11) <> "" Then Tabela("cnac") = Text(11)
If Text(12) <> "" Then Tabela("cprof") = Text(12)
If Combo1(1) <> "" Then Tabela("cestcivil") = Combo1(1)
If Text(13) <> "" Then Tabela("ccpf") = Text(13)
If Text(14) <> "" Then Tabela("crg") = Text(14)
If Text(15) <> "" Then Tabela("cend") = Text(15)
If Text(16) <> "" Then Tabela("cbairro") = Text(16)
If Text(17) <> "" Then Tabela("ccidade") = Text(17)
If Text(18) <> "" Then Tabela("cuf") = Text(18)
If Text(19) <> "" Then Tabela("Ccep") = Text(19)
If Text(20) <> "" Then Tabela("vconjuge") = Text(20)
If Text(21) <> "" Then Tabela("vnacconjuge") = Text(21)
If Text(22) <> "" Then Tabela("vprofconjuge") = Text(22)
If Text(23) <> "" Then Tabela("vrgconjuge") = Text(23)
If Text(24) <> "" Then Tabela("vcpfconjuge") = Text(24)
If Text(25) <> "" Then Tabela("cconjuge") = Text(25)
If Text(26) <> "" Then Tabela("cnacconjuge") = Text(26)
If Text(27) <> "" Then Tabela("cprofconjuge") = Text(27)
If Text(28) <> "" Then Tabela("crgconjuge") = Text(28)
If Text(29) <> "" Then Tabela("ccpfconjuge") = Text(29)
If Text(31) <> "" Then Tabela("Desc") = Text(31)
If Text(32) <> "" Then Tabela("Neg") = Text(32)

Tabela.Update

End Function

Private Function GravarB()

Tabelab.AddNew

If Text1 <> "" Then Tabelab("codigo") = Text1
If Text(0) <> "" Then Tabelab("vendedor") = Text(0)
If Text(1) <> "" Then Tabelab("nacional") = Text(1)
If Text(2) <> "" Then Tabelab("prof") = Text(2)
If Combo1(0) <> "" Then Tabelab("estcivil") = Combo1(0)
If Text(3) <> "" Then Tabelab("rg") = Text(3)
If Text(4) <> "" Then Tabelab("cpf") = Text(4)
If Text(5) <> "" Then Tabelab("vend") = Text(5)
If Text(6) <> "" Then Tabelab("vbairro") = Text(6)
If Text(7) <> "" Then Tabelab("vcidade") = Text(7)
If Text(8) <> "" Then Tabelab("vuf") = Text(8)
If Text(9) <> "" Then Tabelab("Vcep") = Text(9)
If Text(10) <> "" Then Tabelab("comprador") = Text(10)
If Text(11) <> "" Then Tabelab("cnac") = Text(11)
If Text(12) <> "" Then Tabelab("cprof") = Text(12)
If Combo1(1) <> "" Then Tabelab("cestcivil") = Combo1(1)
If Text(13) <> "" Then Tabelab("ccpf") = Text(13)
If Text(14) <> "" Then Tabelab("crg") = Text(14)
If Text(15) <> "" Then Tabelab("cend") = Text(15)
If Text(16) <> "" Then Tabelab("cbairro") = Text(16)
If Text(17) <> "" Then Tabelab("ccidade") = Text(17)
If Text(18) <> "" Then Tabelab("cuf") = Text(18)
If Text(19) <> "" Then Tabelab("Ccep") = Text(19)
If Text(20) <> "" Then Tabelab("vconjuge") = Text(20)
If Text(21) <> "" Then Tabelab("vnacconjuge") = Text(21)
If Text(22) <> "" Then Tabelab("vprofconjuge") = Text(22)
If Text(23) <> "" Then Tabelab("vrgconjuge") = Text(23)
If Text(24) <> "" Then Tabelab("vcpfconjuge") = Text(24)
If Text(25) <> "" Then Tabelab("cconjuge") = Text(25)
If Text(26) <> "" Then Tabelab("cnacconjuge") = Text(26)
If Text(27) <> "" Then Tabelab("cprofconjuge") = Text(27)
If Text(28) <> "" Then Tabelab("crgconjuge") = Text(28)
If Text(29) <> "" Then Tabelab("ccpfconjuge") = Text(29)
If Text(31) <> "" Then Tabelab("Desc") = Text(31)
If Text(32) <> "" Then Tabelab("Neg") = Text(32)

Tabelab.Update

End Function

Private Function Alterar()

Tabela.Edit
If Text(0) <> "" Then Tabela("vendedor") = Text(0)
If Text(1) <> "" Then Tabela("nacional") = Text(1)
If Text(2) <> "" Then Tabela("prof") = Text(2)
If Combo1(0) <> "" Then Tabela("estcivil") = Combo1(0)
If Text(3) <> "" Then Tabela("rg") = Text(3)
If Text(4) <> "" Then Tabela("cpf") = Text(4)
If Text(5) <> "" Then Tabela("vend") = Text(5)
If Text(6) <> "" Then Tabela("vbairro") = Text(6)
If Text(7) <> "" Then Tabela("vcidade") = Text(7)
If Text(8) <> "" Then Tabela("vuf") = Text(8)
If Text(9) <> "" Then Tabela("Vcep") = Text(9)
If Text(10) <> "" Then Tabela("comprador") = Text(10)
If Text(11) <> "" Then Tabela("cnac") = Text(11)
If Text(12) <> "" Then Tabela("cprof") = Text(12)
If Combo1(1) <> "" Then Tabela("cestcivil") = Combo1(1)
If Text(13) <> "" Then Tabela("ccpf") = Text(13)
If Text(14) <> "" Then Tabela("crg") = Text(14)
If Text(15) <> "" Then Tabela("cend") = Text(15)
If Text(16) <> "" Then Tabela("cbairro") = Text(16)
If Text(17) <> "" Then Tabela("ccidade") = Text(17)
If Text(18) <> "" Then Tabela("cuf") = Text(18)
If Text(19) <> "" Then Tabela("Ccep") = Text(19)
If Text(20) <> "" Then Tabela("vconjuge") = Text(20)
If Text(21) <> "" Then Tabela("vnacconjuge") = Text(21)
If Text(22) <> "" Then Tabela("vprofconjuge") = Text(22)
If Text(23) <> "" Then Tabela("vrgconjuge") = Text(23)
If Text(24) <> "" Then Tabela("vcpfconjuge") = Text(24)
If Text(25) <> "" Then Tabela("cconjuge") = Text(25)
If Text(26) <> "" Then Tabela("cnacconjuge") = Text(26)
If Text(27) <> "" Then Tabela("cprofconjuge") = Text(27)
If Text(28) <> "" Then Tabela("crgconjuge") = Text(28)
If Text(29) <> "" Then Tabela("ccpfconjuge") = Text(29)
If Text(31) <> "" Then Tabela("Desc") = Text(31)
If Text(32) <> "" Then Tabela("Neg") = Text(32)

Tabela.Update

End Function

Private Function AlterarB()

Tabelab.Edit
If Text(0) <> "" Then Tabelab("vendedor") = Text(0)
If Text(1) <> "" Then Tabelab("nacional") = Text(1)
If Text(2) <> "" Then Tabelab("prof") = Text(2)
If Combo1(0) <> "" Then Tabelab("estcivil") = Combo1(0)
If Text(3) <> "" Then Tabelab("rg") = Text(3)
If Text(4) <> "" Then Tabelab("cpf") = Text(4)
If Text(5) <> "" Then Tabelab("vend") = Text(5)
If Text(6) <> "" Then Tabelab("vbairro") = Text(6)
If Text(7) <> "" Then Tabelab("vcidade") = Text(7)
If Text(8) <> "" Then Tabelab("vuf") = Text(8)
If Text(9) <> "" Then Tabelab("Vcep") = Text(9)
If Text(10) <> "" Then Tabelab("comprador") = Text(10)
If Text(11) <> "" Then Tabelab("cnac") = Text(11)
If Text(12) <> "" Then Tabelab("cprof") = Text(12)
If Combo1(1) <> "" Then Tabelab("cestcivil") = Combo1(1)
If Text(13) <> "" Then Tabelab("ccpf") = Text(13)
If Text(14) <> "" Then Tabelab("crg") = Text(14)
If Text(15) <> "" Then Tabelab("cend") = Text(15)
If Text(16) <> "" Then Tabelab("cbairro") = Text(16)
If Text(17) <> "" Then Tabelab("ccidade") = Text(17)
If Text(18) <> "" Then Tabelab("cuf") = Text(18)
If Text(19) <> "" Then Tabelab("Ccep") = Text(19)
If Text(20) <> "" Then Tabelab("vconjuge") = Text(20)
If Text(21) <> "" Then Tabelab("vnacconjuge") = Text(21)
If Text(22) <> "" Then Tabelab("vprofconjuge") = Text(22)
If Text(23) <> "" Then Tabelab("vrgconjuge") = Text(23)
If Text(24) <> "" Then Tabelab("vcpfconjuge") = Text(24)
If Text(25) <> "" Then Tabelab("cconjuge") = Text(25)
If Text(26) <> "" Then Tabelab("cnacconjuge") = Text(26)
If Text(27) <> "" Then Tabelab("cprofconjuge") = Text(27)
If Text(28) <> "" Then Tabelab("crgconjuge") = Text(28)
If Text(29) <> "" Then Tabelab("ccpfconjuge") = Text(29)
If Text(31) <> "" Then Tabelab("Desc") = Text(31)
If Text(32) <> "" Then Tabelab("Neg") = Text(32)

Tabela.Update

End Function

Private Function CarregaDadosA()

If Tabela("codigo") <> "" Then Text3 = "KCN" & Format(Tabela("codigo"), "000")
If Tabela("vendedor") <> "" Then Text(0) = Tabela("vendedor")
If Tabela("nacional") <> "" Then Text(1) = Tabela("nacional")
If Tabela("prof") <> "" Then Text(2) = Tabela("prof")
If Tabela("estcivil") <> "" Then Combo1(0) = Tabela("estcivil")
If Tabela("rg") <> "" Then Text(3) = Tabela("rg")
If Tabela("cpf") <> "" Then Text(4) = Tabela("cpf")
If Tabela("vend") <> "" Then Text(5) = Tabela("vend")
If Tabela("vbairro") <> "" Then Text(6) = Tabela("vbairro")
If Tabela("vcidade") <> "" Then Text(7) = Tabela("vcidade")
If Tabela("vuf") <> "" Then Text(8) = Tabela("vuf")
If Tabela("Vcep") <> "" Then Text(9) = Tabela("Vcep")
If Tabela("comprador") <> "" Then Text(10) = Tabela("comprador")
If Tabela("cnac") <> "" Then Text(11) = Tabela("cnac")
If Tabela("cprof") <> "" Then Text(12) = Tabela("cprof")
If Tabela("cestcivil") <> "" Then Combo1(1) = Tabela("cestcivil")
If Tabela("crg") <> "" Then Text(13) = Tabela("crg")
If Tabela("ccpf") <> "" Then Text(14) = Tabela("ccpf")
If Tabela("cend") <> "" Then Text(15) = Tabela("cend")
If Tabela("cbairro") <> "" Then Text(16) = Tabela("cbairro")
If Tabela("ccidade") <> "" Then Text(17) = Tabela("ccidade")
If Tabela("cuf") <> "" Then Text(18) = Tabela("cuf")
If Tabela("Ccep") <> "" Then Text(19) = Tabela("Ccep")
If Tabela("vconjuge") <> "" Then Text(20) = Tabela("vconjuge")
If Tabela("vnacconjuge") <> "" Then Text(21) = Tabela("vnacconjuge")
If Tabela("vprofconjuge") <> "" Then Text(22) = Tabela("vprofconjuge")
If Tabela("vrgconjuge") <> "" Then Text(23) = Tabela("vrgconjuge")
If Tabela("vcpfconjuge") <> "" Then Text(24) = Tabela("vcpfconjuge")
If Tabela("cconjuge") <> "" Then Text(25) = Tabela("cconjuge")
If Tabela("cnacconjuge") <> "" Then Text(26) = Tabela("cnacconjuge")
If Tabela("cprofconjuge") <> "" Then Text(27) = Tabela("cprofconjuge")
If Tabela("crgconjuge") <> "" Then Text(28) = Tabela("crgconjuge")
If Tabela("ccpfconjuge") <> "" Then Text(29) = Tabela("ccpfconjuge")
If Tabela("Desc") <> "" Then Text(31) = Tabela("Desc")
If Tabela("Neg") <> "" Then Text(32) = Tabela("Neg")

End Function

Private Function CarregaDadosB()

If Tabelab("codigo") <> "" Then Text3 = "KCA" & Format(Tabelab("codigo"), "000")
If Tabelab("vendedor") <> "" Then Text(0) = Tabelab("vendedor")
If Tabelab("nacional") <> "" Then Text(1) = Tabelab("nacional")
If Tabelab("prof") <> "" Then Text(2) = Tabelab("prof")
If Tabelab("estcivil") <> "" Then Combo1(0) = Tabelab("estcivil")
If Tabelab("rg") <> "" Then Text(3) = Tabelab("rg")
If Tabelab("cpf") <> "" Then Text(4) = Tabelab("cpf")
If Tabelab("vend") <> "" Then Text(5) = Tabelab("vend")
If Tabelab("vbairro") <> "" Then Text(6) = Tabelab("vbairro")
If Tabelab("vcidade") <> "" Then Text(7) = Tabelab("vcidade")
If Tabelab("vuf") <> "" Then Text(8) = Tabelab("vuf")
If Tabelab("Vcep") <> "" Then Text(9) = Tabelab("Vcep")
If Tabelab("comprador") <> "" Then Text(10) = Tabelab("comprador")
If Tabelab("cnac") <> "" Then Text(11) = Tabelab("cnac")
If Tabelab("cprof") <> "" Then Text(12) = Tabelab("cprof")
If Tabelab("cestcivil") <> "" Then Combo1(1) = Tabelab("cestcivil")
If Tabelab("crg") <> "" Then Text(13) = Tabelab("crg")
If Tabelab("ccpf") <> "" Then Text(14) = Tabelab("ccpf")
If Tabelab("cend") <> "" Then Text(15) = Tabelab("cend")
If Tabelab("cbairro") <> "" Then Text(16) = Tabelab("cbairro")
If Tabelab("ccidade") <> "" Then Text(17) = Tabelab("ccidade")
If Tabelab("cuf") <> "" Then Text(18) = Tabelab("cuf")
If Tabelab("Ccep") <> "" Then Text(19) = Tabelab("Ccep")
If Tabelab("vconjuge") <> "" Then Text(20) = Tabelab("vconjuge")
If Tabelab("vnacconjuge") <> "" Then Text(21) = Tabelab("vnacconjuge")
If Tabelab("vprofconjuge") <> "" Then Text(22) = Tabelab("vprofconjuge")
If Tabelab("vrgconjuge") <> "" Then Text(23) = Tabelab("vrgconjuge")
If Tabelab("vcpfconjuge") <> "" Then Text(24) = Tabelab("vcpfconjuge")
If Tabelab("cconjuge") <> "" Then Text(25) = Tabelab("cconjuge")
If Tabelab("cnacconjuge") <> "" Then Text(26) = Tabelab("cnacconjuge")
If Tabelab("cprofconjuge") <> "" Then Text(27) = Tabelab("cprofconjuge")
If Tabelab("crgconjuge") <> "" Then Text(28) = Tabelab("crgconjuge")
If Tabelab("ccpfconjuge") <> "" Then Text(29) = Tabelab("ccpfconjuge")
If Tabelab("Desc") <> "" Then Text(31) = Tabelab("Desc")
If Tabelab("Neg") <> "" Then Text(32) = Tabelab("Neg")

End Function

Private Sub Text3_Change()
If Len(Text3) = 0 Then
    FechaTab0
    FechaTab1
    FechaTab2
    FechaTab3
Else
    AbreTab0
    AbreTab1
    AbreTab2
    AbreTab3
End If
End Sub

Private Sub Text3_GotFocus()
Text3.BackColor = &HFFFF&
End Sub

Private Sub Text3_LostFocus()
Text3.BackColor = &H8000000B
Text3.Enabled = False
End Sub
