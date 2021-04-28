VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLoc 
   BorderStyle     =   0  'None
   Caption         =   "Super Imob - Cadastro Contrato de Locação"
   ClientHeight    =   6165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Opções:"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   89
      Top             =   600
      Width           =   3615
      Begin VB.OptionButton Option5 
         Caption         =   "Novo Contrato"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   91
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Cadastro de Contrato"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1680
         TabIndex        =   90
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Digite o ano do contrato:"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3840
      TabIndex        =   88
      Top             =   600
      Width           =   2055
   End
   Begin VB.Frame Frame5 
      Caption         =   "Código Gerado:"
      Height          =   735
      Left            =   7200
      TabIndex        =   85
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
         TabIndex        =   86
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   240
         TabIndex        =   87
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.CommandButton Command3 
      Enabled         =   0   'False
      Height          =   615
      Left            =   240
      Picture         =   "frmLoc.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   84
      ToolTipText     =   "Novo Contrato"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Enabled         =   0   'False
      Height          =   615
      Left            =   1560
      Picture         =   "frmLoc.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   83
      ToolTipText     =   "Alterar Contrato"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Enabled         =   0   'False
      Height          =   615
      Left            =   2880
      Picture         =   "frmLoc.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   82
      ToolTipText     =   "Excluir Contrato"
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Enabled         =   0   'False
      Height          =   615
      Left            =   4080
      Picture         =   "frmLoc.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   81
      ToolTipText     =   "Gravar Contrato"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Enabled         =   0   'False
      Height          =   615
      Left            =   5400
      Picture         =   "frmLoc.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   80
      ToolTipText     =   "Limpar Formulário"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Height          =   615
      Left            =   8040
      Picture         =   "frmLoc.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   79
      ToolTipText     =   "Sair"
      Top             =   5280
      Width           =   1095
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
   Begin VB.Frame Frame9 
      Caption         =   "Número:"
      Height          =   735
      Left            =   6000
      TabIndex        =   0
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
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
   Begin TabDlg.SSTab ssPainel 
      Height          =   3375
      Left            =   120
      TabIndex        =   3
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
      TabPicture(0)   =   "frmLoc.frx":198C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Comprador"
      TabPicture(1)   =   "frmLoc.frx":19A8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Conjuge"
      TabPicture(2)   =   "frmLoc.frx":19C4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "Frame2"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Informações Adicionais"
      TabPicture(3)   =   "frmLoc.frx":19E0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "Dados Vendedor:"
         Enabled         =   0   'False
         Height          =   2295
         Index           =   0
         Left            =   240
         TabIndex        =   55
         Top             =   600
         Width           =   8655
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   8
            Left            =   4440
            MaxLength       =   2
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   9
            Left            =   6600
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   7
            Left            =   1080
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   1800
            Width           =   2775
         End
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   6
            Left            =   6000
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   5
            Left            =   1080
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   1440
            Width           =   4215
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   1080
            Width           =   2295
         End
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   6600
            Style           =   2  'Dropdown List
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   6600
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   360
            Width           =   4575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Cpf:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5040
            TabIndex        =   58
            Top             =   1080
            Width           =   615
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Cnpj:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5880
            TabIndex        =   57
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   6600
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Uf:"
            Height          =   195
            Index           =   0
            Left            =   4080
            TabIndex        =   78
            Top             =   1800
            Width           =   210
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Cep:"
            Height          =   195
            Index           =   0
            Left            =   6120
            TabIndex        =   77
            Top             =   1800
            Width           =   330
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   76
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Index           =   0
            Left            =   5400
            TabIndex        =   75
            Top             =   1440
            Width           =   450
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   74
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rg:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   73
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil:"
            Height          =   195
            Index           =   0
            Left            =   5400
            TabIndex        =   72
            Top             =   720
            Width           =   870
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   71
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Index           =   0
            Left            =   5400
            TabIndex        =   70
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Width           =   465
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados Comprador:"
         Height          =   2295
         Index           =   1
         Left            =   -74760
         TabIndex        =   31
         Top             =   600
         Width           =   8655
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   18
            Left            =   4440
            MaxLength       =   2
            TabIndex        =   44
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   19
            Left            =   6600
            TabIndex        =   43
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   17
            Left            =   1080
            TabIndex        =   42
            Top             =   1800
            Width           =   2775
         End
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   16
            Left            =   6000
            TabIndex        =   41
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   15
            Left            =   1080
            TabIndex        =   40
            Top             =   1440
            Width           =   4215
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   13
            Left            =   1080
            TabIndex        =   39
            Top             =   1080
            Width           =   2295
         End
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   6600
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   12
            Left            =   1080
            TabIndex        =   37
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   11
            Left            =   6600
            TabIndex        =   36
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   10
            Left            =   720
            TabIndex        =   35
            Top             =   360
            Width           =   4575
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Cpf:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5040
            TabIndex        =   34
            Top             =   1080
            Width           =   615
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Cnpj:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5880
            TabIndex        =   33
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   14
            Left            =   6600
            TabIndex        =   32
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Uf:"
            Height          =   195
            Index           =   1
            Left            =   4080
            TabIndex        =   54
            Top             =   1800
            Width           =   210
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
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   52
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Index           =   1
            Left            =   5400
            TabIndex        =   51
            Top             =   1440
            Width           =   450
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   50
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rg:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   49
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil:"
            Height          =   195
            Index           =   1
            Left            =   5400
            TabIndex        =   48
            Top             =   720
            Width           =   870
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   47
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Index           =   1
            Left            =   5400
            TabIndex        =   46
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   465
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Conjuge Vendedor:"
         Enabled         =   0   'False
         Height          =   1215
         Left            =   -74760
         TabIndex        =   20
         Top             =   480
         Width           =   8655
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   20
            Left            =   720
            TabIndex        =   25
            Top             =   360
            Width           =   4575
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   21
            Left            =   6600
            TabIndex        =   24
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   22
            Left            =   960
            TabIndex        =   23
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   23
            Left            =   3720
            TabIndex        =   22
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   24
            Left            =   6480
            TabIndex        =   21
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Left            =   5400
            TabIndex        =   29
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Rg:"
            Height          =   195
            Left            =   3360
            TabIndex        =   27
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Cpf:"
            Height          =   195
            Left            =   6120
            TabIndex        =   26
            Top             =   720
            Width           =   285
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Conjuge Comprador:"
         Enabled         =   0   'False
         Height          =   1215
         Left            =   -74760
         TabIndex        =   9
         Top             =   1800
         Width           =   8655
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   29
            Left            =   6480
            TabIndex        =   14
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   28
            Left            =   3720
            TabIndex        =   13
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   27
            Left            =   960
            TabIndex        =   12
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   26
            Left            =   6600
            TabIndex        =   11
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   25
            Left            =   720
            TabIndex        =   10
            Top             =   360
            Width           =   4575
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Cpf:"
            Height          =   195
            Left            =   6120
            TabIndex        =   19
            Top             =   720
            Width           =   285
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Rg:"
            Height          =   195
            Left            =   3360
            TabIndex        =   18
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Left            =   5400
            TabIndex        =   16
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   465
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Informações do Imóvel à Venda:"
         Height          =   2655
         Left            =   -74760
         TabIndex        =   4
         Top             =   480
         Width           =   8655
         Begin VB.TextBox Text 
            Height          =   735
            Index           =   31
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   600
            Width           =   8415
         End
         Begin VB.TextBox Text 
            Height          =   735
            Index           =   32
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   1800
            Width           =   8415
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   765
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Informações da Negociação:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   1560
            Width           =   2055
         End
      End
   End
   Begin VB.Frame Frame8 
      Height          =   975
      Left            =   120
      TabIndex        =   92
      Top             =   5040
      Width           =   9135
      Begin VB.CommandButton Command9 
         Enabled         =   0   'False
         Height          =   615
         Left            =   6600
         Picture         =   "frmLoc.frx":19FC
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Buscar Contrato"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CONTRATO DE LOCAÇÃO"
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
      TabIndex        =   96
      Top             =   120
      Width           =   2310
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   0
      X2              =   9360
      Y1              =   360
      Y2              =   360
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
      TabIndex        =   94
      Top             =   120
      Width           =   885
   End
End
Attribute VB_Name = "frmLoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
