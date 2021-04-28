VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmCompraVenda 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Super Imob - Compra e Venda"
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programa Imobiliária\Dados\Bdimobiliaria.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Prestacao"
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programa Imobiliária\Dados\Bdimobiliaria.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "VContrato"
      Top             =   360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Número:"
      Height          =   735
      Left            =   6000
      TabIndex        =   82
      Top             =   600
      Width           =   1095
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         DataField       =   "Codigo"
         DataSource      =   "Data1"
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
         TabIndex        =   83
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
      Height          =   3495
      Left            =   120
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1440
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   6165
      _Version        =   393216
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Vendedor"
      TabPicture(0)   =   "frmCompraVenda1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Comprador"
      TabPicture(1)   =   "frmCompraVenda1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Conjuges"
      TabPicture(2)   =   "frmCompraVenda1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Informações"
      TabPicture(3)   =   "frmCompraVenda1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Negociação"
      TabPicture(4)   =   "frmCompraVenda1.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame10"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Buscar"
      TabPicture(5)   =   "frmCompraVenda1.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame13"
      Tab(5).Control(1)=   "Frame14"
      Tab(5).ControlCount=   2
      Begin VB.Frame Frame14 
         Caption         =   "Digite para buscar:"
         Height          =   615
         Left            =   -74760
         TabIndex        =   142
         Top             =   480
         Width           =   8655
         Begin VB.TextBox Text95 
            BackColor       =   &H00C0FFC0&
            Height          =   315
            Left            =   120
            TabIndex        =   147
            Top             =   240
            Width           =   2895
         End
         Begin VB.OptionButton Option11 
            Caption         =   "Código"
            Height          =   195
            Left            =   5040
            TabIndex        =   146
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option10 
            Caption         =   "Vendedor"
            Height          =   195
            Left            =   6120
            TabIndex        =   145
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Comprador"
            Height          =   195
            Left            =   7440
            TabIndex        =   144
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command11 
            Caption         =   ">>> Busc&ar >>>"
            Height          =   300
            Left            =   3120
            TabIndex        =   143
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame13 
         Height          =   2175
         Left            =   -74760
         TabIndex        =   140
         Top             =   1080
         Width           =   8655
         Begin MSDBGrid.DBGrid DBGrid2 
            Bindings        =   "frmCompraVenda1.frx":00A8
            Height          =   1815
            Left            =   120
            OleObjectBlob   =   "frmCompraVenda1.frx":00BC
            TabIndex        =   141
            Top             =   240
            Width           =   8415
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Informações sobre a negociação:"
         Height          =   2775
         Left            =   240
         TabIndex        =   123
         Top             =   480
         Width           =   8655
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            DataField       =   "Entr"
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
            Index           =   54
            Left            =   3000
            TabIndex        =   133
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Gerar Parcelas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   7200
            Style           =   1  'Graphical
            TabIndex        =   138
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            DataField       =   "Prest"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   53
            Left            =   5880
            TabIndex        =   137
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            DataField       =   "DataNeg"
            DataSource      =   "Data1"
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   52
            Left            =   5880
            TabIndex        =   136
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            DataField       =   "ValorVenda"
            DataSource      =   "Data1"
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
            Index           =   51
            Left            =   3000
            TabIndex        =   135
            Top             =   720
            Width           =   1335
         End
         Begin VB.Frame Frame12 
            Caption         =   "Vencimentos Gerados:"
            Height          =   1695
            Left            =   120
            TabIndex        =   125
            Top             =   960
            Width           =   8415
            Begin MSDBGrid.DBGrid DBGrid1 
               Bindings        =   "frmCompraVenda1.frx":0DFF
               Height          =   1215
               Left            =   120
               OleObjectBlob   =   "frmCompraVenda1.frx":0E13
               TabIndex        =   126
               Top             =   240
               Width           =   8175
            End
            Begin VB.Label Label48 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Height          =   195
               Left            =   2520
               TabIndex        =   139
               Top             =   0
               Visible         =   0   'False
               Width           =   45
            End
         End
         Begin VB.Frame Frame11 
            Height          =   615
            Left            =   120
            TabIndex        =   124
            Top             =   240
            Width           =   2415
            Begin VB.OptionButton Option2 
               Caption         =   "À Prazo"
               Enabled         =   0   'False
               Height          =   195
               Left            =   1320
               TabIndex        =   152
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton Option1 
               Caption         =   "À Vista"
               Enabled         =   0   'False
               Height          =   195
               Left            =   120
               TabIndex        =   151
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "R$:"
            Height          =   195
            Left            =   2640
            TabIndex        =   134
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Ano/Negociação:"
            Height          =   195
            Left            =   4440
            TabIndex        =   129
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Qtde. Prestação:"
            Height          =   195
            Left            =   4440
            TabIndex        =   128
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Entr:"
            Height          =   195
            Left            =   2640
            TabIndex        =   127
            Top             =   240
            Width           =   330
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Informações do Contrato:"
         Height          =   2895
         Left            =   -74760
         TabIndex        =   81
         Top             =   480
         Width           =   8655
         Begin VB.TextBox Text 
            DataField       =   "Obs"
            DataSource      =   "Data1"
            Height          =   495
            Index           =   55
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   154
            Top             =   2280
            Width           =   7335
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "<< Adicion&ar Imóvel >>"
            Height          =   375
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   130
            Top             =   120
            Width           =   2655
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   31
            Left            =   7080
            TabIndex        =   121
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "AreaU"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   48
            Left            =   7680
            TabIndex        =   119
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "AreaC"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   47
            Left            =   6480
            TabIndex        =   117
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "AreaT"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   46
            Left            =   5160
            TabIndex        =   116
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "LadoE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   45
            Left            =   3960
            TabIndex        =   113
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "LadoD"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   44
            Left            =   2760
            TabIndex        =   111
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Fundos"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   43
            Left            =   1560
            TabIndex        =   109
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "TestP"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   42
            Left            =   120
            TabIndex        =   107
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Cep"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   41
            Left            =   6960
            TabIndex        =   105
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Uf"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   40
            Left            =   6120
            TabIndex        =   104
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox Text 
            DataField       =   "Cidade"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   39
            Left            =   3120
            TabIndex        =   101
            Top             =   1320
            Width           =   2895
         End
         Begin VB.TextBox Text 
            DataField       =   "Bairro"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   38
            Left            =   120
            TabIndex        =   100
            Top             =   1320
            Width           =   2895
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Cond"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   37
            Left            =   7680
            TabIndex        =   98
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Andar"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   36
            Left            =   6960
            TabIndex        =   96
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "elevador"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   35
            Left            =   6240
            TabIndex        =   94
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Bloco"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   34
            Left            =   5640
            TabIndex        =   92
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Apto"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   33
            Left            =   5040
            TabIndex        =   90
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox Text 
            DataField       =   "EndImovel"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   32
            Left            =   120
            TabIndex        =   88
            Top             =   720
            Width           =   4815
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Informações:"
            Height          =   195
            Left            =   120
            TabIndex        =   153
            Top             =   2400
            Width           =   915
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Valor do Imóvel R$:"
            Height          =   195
            Left            =   5640
            TabIndex        =   122
            Top             =   120
            Width           =   1395
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Área Útil:"
            Height          =   195
            Left            =   7680
            TabIndex        =   120
            Top             =   1680
            Width           =   645
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Área Constr.:"
            Height          =   195
            Left            =   6480
            TabIndex        =   118
            Top             =   1680
            Width           =   915
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Área Terreno:"
            Height          =   195
            Left            =   5160
            TabIndex        =   115
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Lado Esquerdo:"
            Height          =   195
            Left            =   3960
            TabIndex        =   114
            Top             =   1680
            Width           =   1125
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Lado direito:"
            Height          =   195
            Left            =   2760
            TabIndex        =   112
            Top             =   1680
            Width           =   870
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Fundos:"
            Height          =   195
            Left            =   1560
            TabIndex        =   110
            Top             =   1680
            Width           =   570
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Testada Principal:"
            Height          =   195
            Left            =   120
            TabIndex        =   108
            Top             =   1680
            Width           =   1275
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Cep:"
            Height          =   195
            Left            =   6960
            TabIndex        =   106
            Top             =   1080
            Width           =   330
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            Height          =   195
            Left            =   6120
            TabIndex        =   103
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Left            =   3120
            TabIndex        =   102
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Left            =   120
            TabIndex        =   99
            Top             =   1080
            Width           =   450
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Condomínio:"
            Height          =   195
            Left            =   7680
            TabIndex        =   97
            Top             =   480
            Width           =   900
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Andar:"
            Height          =   195
            Left            =   6960
            TabIndex        =   95
            Top             =   480
            Width           =   465
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Elevador:"
            Height          =   195
            Left            =   6240
            TabIndex        =   93
            Top             =   480
            Width           =   675
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Bloco:"
            Height          =   195
            Left            =   5640
            TabIndex        =   91
            Top             =   480
            Width           =   450
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Apto:"
            Height          =   195
            Left            =   5040
            TabIndex        =   89
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Left            =   120
            TabIndex        =   87
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Conjuge Comprador:"
         Enabled         =   0   'False
         Height          =   1215
         Left            =   -74760
         TabIndex        =   70
         Top             =   2040
         Width           =   8655
         Begin VB.TextBox Text 
            DataField       =   "Cconjuge"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   25
            Left            =   720
            TabIndex        =   75
            Top             =   360
            Width           =   4575
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Cnacconjuge"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   26
            Left            =   6600
            TabIndex        =   74
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Cprofconjuge"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   27
            Left            =   960
            TabIndex        =   73
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Crgconjuge"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   28
            Left            =   3720
            TabIndex        =   72
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Ccpfconjuge"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   29
            Left            =   6480
            TabIndex        =   71
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Left            =   5400
            TabIndex        =   79
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Left            =   120
            TabIndex        =   78
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Rg:"
            Height          =   195
            Left            =   3360
            TabIndex        =   77
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Cpf:"
            Height          =   195
            Left            =   6120
            TabIndex        =   76
            Top             =   720
            Width           =   285
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Conjuge Vendedor:"
         Enabled         =   0   'False
         Height          =   1215
         Left            =   -74760
         TabIndex        =   59
         Top             =   600
         Width           =   8655
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Vcpfconjuge"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   24
            Left            =   6480
            TabIndex        =   69
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Vrgconjuge"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   23
            Left            =   3720
            TabIndex        =   67
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Vprofconjuge"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   22
            Left            =   960
            TabIndex        =   64
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Vnacconjuge"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   21
            Left            =   6600
            TabIndex        =   62
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text 
            DataField       =   "Vconjuge"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   20
            Left            =   720
            TabIndex        =   60
            Top             =   360
            Width           =   4575
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Cpf:"
            Height          =   195
            Left            =   6120
            TabIndex        =   68
            Top             =   720
            Width           =   285
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Rg:"
            Height          =   195
            Left            =   3360
            TabIndex        =   66
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Left            =   5400
            TabIndex        =   63
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Width           =   465
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados Comprador:"
         Height          =   2655
         Index           =   1
         Left            =   -74760
         TabIndex        =   15
         Top             =   600
         Width           =   8655
         Begin VB.CommandButton Command7 
            BackColor       =   &H00C0FFFF&
            Caption         =   "<< Adicion&ar Cliente >>"
            Height          =   435
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   132
            Top             =   120
            Width           =   2895
         End
         Begin VB.TextBox Text 
            DataField       =   "Cestcivil"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   50
            Left            =   6600
            TabIndex        =   86
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Ccpf"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   14
            Left            =   6600
            TabIndex        =   20
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox Text 
            DataField       =   "Comprador"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   10
            Left            =   720
            TabIndex        =   16
            Top             =   720
            Width           =   4575
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "CNac"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   11
            Left            =   6600
            TabIndex        =   17
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Cprof"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   12
            Left            =   1080
            TabIndex        =   18
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Crg"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   13
            Left            =   1080
            TabIndex        =   19
            Top             =   1440
            Width           =   2295
         End
         Begin VB.TextBox Text 
            DataField       =   "Cend"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   15
            Left            =   1080
            TabIndex        =   21
            Top             =   1800
            Width           =   4215
         End
         Begin VB.TextBox Text 
            DataField       =   "Cbairro"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   16
            Left            =   6000
            TabIndex        =   22
            Top             =   1800
            Width           =   2535
         End
         Begin VB.TextBox Text 
            DataField       =   "Ccidade"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   17
            Left            =   1080
            TabIndex        =   23
            Top             =   2160
            Width           =   2775
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Ccep"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   19
            Left            =   6600
            TabIndex        =   25
            Top             =   2160
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Cuf"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   18
            Left            =   4440
            MaxLength       =   2
            TabIndex        =   24
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "CPF / CNPJ:"
            Height          =   195
            Left            =   5400
            TabIndex        =   149
            Top             =   1440
            Width           =   915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   58
            Top             =   720
            Width           =   465
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Index           =   1
            Left            =   5400
            TabIndex        =   57
            Top             =   720
            Width           =   1065
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   56
            Top             =   1080
            Width           =   690
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil:"
            Height          =   195
            Index           =   1
            Left            =   5400
            TabIndex        =   55
            Top             =   1080
            Width           =   870
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rg:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   54
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   53
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Index           =   1
            Left            =   5400
            TabIndex        =   52
            Top             =   1800
            Width           =   450
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   51
            Top             =   2160
            Width           =   540
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Cep:"
            Height          =   195
            Index           =   1
            Left            =   6120
            TabIndex        =   50
            Top             =   2160
            Width           =   330
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Uf:"
            Height          =   195
            Index           =   1
            Left            =   4080
            TabIndex        =   49
            Top             =   2160
            Width           =   210
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados Vendedor:"
         Enabled         =   0   'False
         Height          =   2655
         Index           =   0
         Left            =   -74760
         TabIndex        =   4
         Top             =   600
         Width           =   8655
         Begin VB.CommandButton Command10 
            BackColor       =   &H00C0FFFF&
            Caption         =   "<< Adicion&ar Cliente >>"
            Height          =   435
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   131
            Top             =   120
            Width           =   2895
         End
         Begin VB.TextBox Text 
            DataField       =   "EstCivil"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   49
            Left            =   6600
            TabIndex        =   85
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Cpf"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   4
            Left            =   6600
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox Text 
            DataField       =   "Vendedor"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   720
            Width           =   4575
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Nacional"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   1
            Left            =   6600
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Prof"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Rg"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1440
            Width           =   2295
         End
         Begin VB.TextBox Text 
            DataField       =   "VEnd"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   5
            Left            =   1080
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1800
            Width           =   4215
         End
         Begin VB.TextBox Text 
            DataField       =   "Vbairro"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   6
            Left            =   6000
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1800
            Width           =   2535
         End
         Begin VB.TextBox Text 
            DataField       =   "VCidade"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   7
            Left            =   1080
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   2160
            Width           =   2775
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Vcep"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   9
            Left            =   6600
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   2160
            Width           =   1935
         End
         Begin VB.TextBox Text 
            Alignment       =   2  'Center
            DataField       =   "Vuf"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   8
            Left            =   4440
            MaxLength       =   2
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "CPF / CNPJ:"
            Height          =   195
            Left            =   5400
            TabIndex        =   148
            Top             =   1440
            Width           =   915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   48
            Top             =   720
            Width           =   465
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Index           =   0
            Left            =   5400
            TabIndex        =   47
            Top             =   720
            Width           =   1065
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   1080
            Width           =   690
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil:"
            Height          =   195
            Index           =   0
            Left            =   5400
            TabIndex        =   45
            Top             =   1080
            Width           =   870
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rg:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   44
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   43
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Index           =   0
            Left            =   5400
            TabIndex        =   42
            Top             =   1800
            Width           =   450
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   41
            Top             =   2160
            Width           =   540
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Cep:"
            Height          =   195
            Index           =   0
            Left            =   6120
            TabIndex        =   40
            Top             =   2160
            Width           =   330
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Uf:"
            Height          =   195
            Index           =   0
            Left            =   4080
            TabIndex        =   39
            Top             =   2160
            Width           =   210
         End
      End
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000000&
      Height          =   615
      Left            =   8040
      Picture         =   "frmCompraVenda1.frx":1EA6
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Sair"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   615
      Left            =   4920
      Picture         =   "frmCompraVenda1.frx":22E8
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Gravar Contrato"
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   615
      Left            =   3360
      Picture         =   "frmCompraVenda1.frx":272A
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Excluir Contrato"
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   615
      Left            =   1800
      Picture         =   "frmCompraVenda1.frx":2B6C
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Alterar Contrato"
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   615
      Left            =   240
      Picture         =   "frmCompraVenda1.frx":2FAE
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Novo Contrato"
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Código Gerado:"
      Height          =   735
      Left            =   7200
      TabIndex        =   29
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
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Digite o ano do contrato:"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3840
      TabIndex        =   28
      Top             =   600
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Opções:"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   27
      Top             =   600
      Width           =   3615
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cadastro de Contrato"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C0C0&
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
      TabIndex        =   36
      Top             =   4920
      Width           =   9135
      Begin VB.CommandButton Command9 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   615
         Left            =   6360
         Picture         =   "frmCompraVenda1.frx":33F0
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Buscar Contrato"
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Label46 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label46"
      DataField       =   "Pagamento"
      DataSource      =   "Data1"
      Height          =   195
      Left            =   4440
      TabIndex        =   150
      Top             =   5880
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Contratos:"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   7320
      TabIndex        =   84
      Top             =   120
      Width           =   720
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
      BackStyle       =   0  'Transparent
      Caption         =   "CONTRATO DE COMPRA E VENDA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   3315
   End
End
Attribute VB_Name = "frmCompraVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dados As DAO.Database
Dim Tabela As DAO.Recordset
Public Sql As String

Private Sub Command1_Click()
Dim mes As Integer
Dim ANO As Integer
Dim Data As String
Dim Valor As String
Dim Parcela As String
Dim Final As String

Final = DateAdd("m", Text(53).Text, CDate(Text(52).Text))

mes = Format(Text(52), "mm")
ANO = Format(Text(52), "yy")
Valor = Val("0")

For i = 1 To Val(Text(53).Text)
    mes = mes + 1
    Valor = Valor + 1
    If mes > 12 Then
       mes = 1
       ANO = ANO + 1
   End If
    
  dia1 = Format(Text(52), "dd")
  dia = Verifica_dia(dia1, mes)
  Data = dia & "/" & Format(mes, "0") & "/" & Format(ANO, "0000")

  Label48 = Text(51).Text - Text(54).Text
  Parcela = Label48.Caption / Val(Text(53).Text)
  Parcela = Format$(Parcela, " ####,###,##0.00")

  Data2.Recordset.AddNew
  Data2.Recordset.Fields(0) = CLng(Text1)
  Data2.Recordset.Fields(1) = CDate(Data)
  Data2.Recordset.Fields(2) = Text(0)
  Data2.Recordset.Fields(3) = Text(10)
  Data2.Recordset.Fields(4) = Parcela
  Data2.Recordset.Fields(5) = Text(52)
  Data2.Recordset.Fields(6) = Final
  Data2.Recordset.Fields(9) = Val(Valor)
  Data2.Recordset.Update
Next
MsgBox ("Prestações Geradas com sucesso!")

Command1.Enabled = False
End Sub

Private Sub Command10_Click()
frmClientes.Command1.Enabled = True
frmClientes.Show 1
End Sub

Private Sub Command11_Click()
On Error Resume Next

If Option11.Value = True Then
    If Text95.Text = "" Then
        Data1.RecordSource = "SELECT * FROM VCONTRATO"
        Data1.Refresh
        Exit Sub
    End If

    Data1.RecordSource = "SELECT * FROM VCONTRATO WHERE CODIGO Like '" & Text95.Text & "*'"
    Data1.Refresh

ElseIf Option10.Value = True Then
    If Text95.Text = "" Then
        Data1.RecordSource = "SELECT * FROM VCONTRATO"
        Data1.Refresh
        Exit Sub
    End If

    Data1.RecordSource = "SELECT * FROM VCONTRATO WHERE VENDEDOR Like '" & Text95.Text & "*'"
    Data1.Refresh
    
ElseIf Option9.Value = True Then
    If Text95.Text = "" Then
        Data1.RecordSource = "SELECT * FROM VCONTRATO"
        Data1.Refresh
        Exit Sub
    End If

    Data1.RecordSource = "SELECT * FROM VCONTRATO WHERE COMPRADOR Like '" & Text95.Text & "*'"
    Data1.Refresh
End If
End Sub

Private Sub Command2_Click()
frmImoveis.cmdImovel.Enabled = True
frmImoveis.Show 1
End Sub

Private Sub Command3_Click()
Data1.Recordset.AddNew
LimpaCaixas
CorTxtAberto
Text(30).Enabled = True
ssPainel.Tab = 0
ssPainel.Enabled = True
Frame1(0).Enabled = True
Frame3.Enabled = True
Frame4.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
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
Data1.Recordset.Edit
ssPainel.Tab = 0
Frame1(0).Enabled = True
AbreTab0
AbreTab1
AbreTab2
AbreTab3
AbreTab4
CorTxtAberto
Text(0).SetFocus
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = True
Command9.Enabled = False
Command8.Picture = LoadPicture(App.Path & "\Ícones\TRFFC14.ico")
Command8.ToolTipText = "Cancelar Novo"
End Sub

Private Sub Command5_Click()
If MsgBox("Confirma Exclusão do Cliente?  -> " & Data1.Recordset![codigo], vbQuestion + vbYesNo, "Excluir Clientes") = vbYes Then
   Data1.Recordset.Delete
   Data1.Refresh
End If
End Sub

Private Sub Command6_Click()

Data1.UpdateRecord
Data1.Refresh
MsgBox ("Alterações Cadastradas com Sucesso!")
FechaTab0
FechaTab1
FechaTab2
FechaTab3
FechaTab4
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
Command8.Picture = LoadPicture(App.Path & "\Ícones\ARW10NE.ico")

End Sub

Private Sub Command7_Click()
frmClientes.Command1.Enabled = True
frmClientes.Show 1
End Sub

Private Sub Command8_Click()
If Command8.ToolTipText = "Cancelar Novo" Then
LimpaCaixas
Text(30).Enabled = False
Data1.Recordset.CancelUpdate
Data1.Refresh

Option1.Value = False
Option2.Value = False
Option1.Enabled = False
Option2.Enabled = False
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
Command4.Enabled = True
Command5.Enabled = True
Command9.Enabled = True
FechaTab0
FechaTab1
FechaTab2
FechaTab3
FechaTab4
Command7.Enabled = False
ssPainel.Tab = 0
Text3.Enabled = False
Text3.BackColor = &H808080
Else
If MsgBox("Quer sair do Cadastro de Contrato?", vbYesNo, "Sair do Cadastro") = vbYes Then
    Unload frmCompraVenda
    RedefineFormPrincipal
  Else
    Exit Sub
End If
End If
End Sub

Private Sub Command9_Click()
ssPainel.Tab = 5
End Sub

Private Sub DBCombo1_LostFocus()
Label46.Caption = DBCombo1.Text
End Sub

Private Sub DBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
On Error GoTo erro

If ColIndex >= 0 And ColIndex <= 5 Then
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
End Sub

Private Sub Form_Load()

Data1.DatabaseName = App.Path & "\Dados\Bdimobiliaria.MDB"
Data1.RecordSource = "VContrato"

Data2.DatabaseName = App.Path & "\Dados\Bdimobiliaria.MDB"
Data2.RecordSource = "Prestacao"

Set Dados = OpenDatabase(App.Path & "\Dados\Bdimobiliaria.MDB")
Set Tabela = Dados.OpenRecordset("VContrato", dbOpenTable)

If Tabela.RecordCount = 0 Then
    FechaTab0
    FechaTab1
    FechaTab2
    FechaTab3
    FechaTab4
    Text(30).Enabled = False
    Text3.BackColor = &H808080
    Text1.BackColor = &H808080
    Label8.Caption = "Contratos Cadastrados: 000"
    Command3.Enabled = True
    Command9.Enabled = True
ElseIf Tabela.RecordCount > 0 Then
    FechaTab0
    FechaTab1
    FechaTab2
    FechaTab3
    FechaTab4
    Text(30).Enabled = False
    Text3.BackColor = &H808080
    Text1.BackColor = &H808080
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command9.Enabled = True
    Label8.Caption = "Contratos Cadastrados: " & Format(Tabela.RecordCount, "000")
End If
ssPainel.Tab = 0
Text95.BackColor = &HC0FFC0
Dados.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmFundo.Enabled = True
End Sub

Private Sub Label46_Change()
If Label46.Caption = "a vista" Then
    Option1.Value = True
ElseIf Label46.Caption = "a prazo" Then
    Option2.Value = True
End If
End Sub

Private Sub Option1_Click()
Label46.Caption = "a vista"
Text(53).Enabled = True
Text(54).Enabled = True
Text(53) = ""
Text(54) = ""
End Sub

Private Sub Option1_KeyPress(KeyAscii As Integer)
Option2.SetFocus
End Sub

Private Sub Option2_Click()
Label46.Caption = "a prazo"
Text(53).Enabled = True
Text(54).Enabled = True
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

Text(30) = Year(Date)
Text(52) = Year(Date)

End Sub

Private Sub Option6_Click()
If Option6.Value = True Then
Text(30).Enabled = True
Text(30).SetFocus
End If
End Sub

Private Sub Option7_Click()
Text(53).Enabled = False
Text(54).Enabled = False
End Sub

Private Sub Option8_Click()

If Command4.Enabled = False Then
    Label46.Caption = "À Prazo"
    Text(53).Enabled = True
    Text(54).Enabled = True
ElseIf Command4.Enabled = True Then
    Label46.Caption = "À Prazo"
    Text(53).Enabled = False
    Text(54).Enabled = False
End If

End Sub

Private Sub Text_Change(Index As Integer)
On Error Resume Next
Select Case Index
Case 30
If Text(30).Enabled = False Then
    Text(30) = Data1.Recordset.Fields(55)
ElseIf Text(30).Enabled = True Then
    If Len(Text(30)) = 0 Then
    Text3 = ""
    End If

    Set Dados = OpenDatabase(App.Path & "\Dados\Bdimobiliaria.MDB")
    Set Tabela = Dados.OpenRecordset("VContrato", dbOpenTable)
    If Len(Text(30)) = 4 Then
    If Text(30) < Year(Date) - 3 Then
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
                        Text1 = "1"
                        Text3 = "KCA001"
                        Text(0).SetFocus
                    Else
                        Tabela.MoveLast
                        Text1 = Tabela("codigo") + 1
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
                    Text1 = "1"
                    Text3 = "KCN001"
                Else
                    Tabela.MoveLast
                    Text1 = Tabela("codigo") + 1
                    Text3 = Text1
                    Text3 = "KCN" & Format(Text3, "000")
                End If
            End If
        End If
    End If
End If
Case 31
    If IsNumeric(Text(31)) = True Then
        Text(31) = Format$(Text(31), " ####,###,##0.00")
    Else
        Text(31) = "R$ 0,00"
    End If
    Text(51).Text = Text(31).Text
Case 52
    Dim ANO As String
    ANO = Format(Text(52), "yyyy")
    Label51.Caption = "Este contrato foi elaborado em " & ANO
End Select
End Sub

Private Sub text_GotFocus(Index As Integer)
RecebeFoco
Select Case Index
    Case 30
        Text(30).BackColor = &HFFFF&
End Select
End Sub

Private Sub Text_LostFocus(Index As Integer)
PerdeFoco
Select Case Index
Case 30
    Text(30).BackColor = &H80000005
    If Not Len(Text(30)) = 4 Then
    MsgBox ("Digite um ano válido de 4 dígitos!!")
    Text(30).SetFocus
    End If
Case 31
    Text(31).BackColor = &H80000005
    If IsNumeric(Text(31)) = True Then
        Text(31) = Format$(Text(31), " ####,###,##0.00")
    Else
        Text(31) = "R$ 0,00"
    End If
Case 51
    If IsNumeric(Text(51)) = True Then
        Text(51) = Format$(Text(51), " ####,###,##0.00")
    Else
        Text(51) = "R$ 0,00"
    End If
Case 54
    If IsNumeric(Text(54)) = True Then
        Text(54) = Format$(Text(54), " ####,###,##0.00")
    ElseIf Text(54) = "" Then
        Text(54) = "0,00"
    End If
End Select
Text(52) = Text(30)
End Sub

Private Function CorTxtFechado()
Dim TxtBox As Object

For Each TxtBox In Me.Controls
    If TypeOf TxtBox Is TextBox Then
    TxtBox.BackColor = &H8000000B
End If
Next TxtBox

End Function

Private Function CorTxtAberto()
Dim TxtBox As Object

For Each TxtBox In Me.Controls
    If TypeOf TxtBox Is TextBox Then
    TxtBox.BackColor = &H80000005
End If
Next TxtBox

End Function

Private Function RecebeFoco()
Dim TxtBox As Object

For Each TxtBox In Me.Controls
    If TypeOf TxtBox Is TextBox Then
    TxtBox.BackColor = &HFFFF&
    TxtBox = StrConv(TxtBox, vbProperCase)
End If
Next TxtBox

End Function

Private Function PerdeFoco()
Dim TxtBox As Object

For Each TxtBox In Me.Controls
    If TypeOf TxtBox Is TextBox Then
    TxtBox.BackColor = &H80000005
    TxtBox = StrConv(TxtBox, vbUpperCase)
End If
Next TxtBox

End Function

Private Function AbreTab0()

CorTxtAberto
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
Text(49).Enabled = True
Command10.Enabled = True

End Function

Private Function FechaTab0()

CorTxtFechado
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
Text(49).Enabled = False
Command10.Enabled = False

End Function

Private Function AbreTab1()

CorTxtAberto
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
Text(50).Enabled = True
Command7.Enabled = True

End Function

Private Function FechaTab1()

CorTxtFechado
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
Text(50).Enabled = False
Command7.Enabled = False

End Function

Private Function AbreTab2()

CorTxtAberto
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

CorTxtFechado
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

CorTxtAberto
Text(31).Enabled = True
Text(32).Enabled = True
Text(33).Enabled = True
Text(34).Enabled = True
Text(35).Enabled = True
Text(36).Enabled = True
Text(37).Enabled = True
Text(38).Enabled = True
Text(39).Enabled = True
Text(40).Enabled = True
Text(41).Enabled = True
Text(42).Enabled = True
Text(43).Enabled = True
Text(44).Enabled = True
Text(45).Enabled = True
Text(46).Enabled = True
Text(47).Enabled = True
Text(48).Enabled = True
Text(55).Enabled = True

Command2.Enabled = True

End Function

Private Function FechaTab3()

CorTxtFechado
Text(31).Enabled = False
Text(32).Enabled = False
Text(33).Enabled = False
Text(34).Enabled = False
Text(35).Enabled = False
Text(36).Enabled = False
Text(37).Enabled = False
Text(38).Enabled = False
Text(39).Enabled = False
Text(40).Enabled = False
Text(41).Enabled = False
Text(42).Enabled = False
Text(43).Enabled = False
Text(44).Enabled = False
Text(45).Enabled = False
Text(46).Enabled = False
Text(47).Enabled = False
Text(48).Enabled = False
Text(55).Enabled = False

Command2.Enabled = False

End Function

Private Function AbreTab4()

CorTxtAberto
Text(51).Enabled = True
Text(52).Enabled = True
Text(53).Enabled = True
Text(54).Enabled = True

Option5.Enabled = True
Option6.Enabled = True


End Function

Private Function FechaTab4()

CorTxtFechado
Text(51).Enabled = False
Text(52).Enabled = False
Text(53).Enabled = False
Text(54).Enabled = True

Option5.Enabled = False
Option6.Enabled = False
Command1.Enabled = False

End Function

Private Function LimpaCaixas()
Dim TextBox As Object

For Each TextBox In Me.Controls
    If TypeOf TextBox Is TextBox Then
    TextBox = Empty
End If
Next TextBox

End Function

Private Sub Text1_Change()
On Error Resume Next
If Text1.Text = "" Then
    Data2.RecordSource = "SELECT * FROM PRESTACAO"
    Data2.Refresh
    Exit Sub
End If

Data2.RecordSource = "SELECT * FROM PRESTACAO WHERE CODIGO Like '" & Text1.Text & "*'"
Data2.Refresh

If Data2.Recordset("Codigo") = "" Then
    Command1.Enabled = True
ElseIf Data1.Recordset("Codigo") = Text1.Text Then
    Command1.Enabled = False
End If
End Sub

Private Sub Text3_Change()
If Len(Text3) = 0 Then
    FechaTab0
    FechaTab1
    FechaTab2
    FechaTab3
    FechaTab4
ElseIf Len(Text3) > 0 Then
    AbreTab0
    AbreTab1
    AbreTab2
    AbreTab3
    AbreTab4
End If
End Sub

Private Sub Text3_GotFocus()
Text3.BackColor = &HFFFF&
End Sub

Private Sub Text3_LostFocus()
Text3.BackColor = &H8000000B
Text3.Enabled = False
End Sub

Public Function Verifica_dia(dia, mes)
Dim diasDoMes As Variant

dia = Val(dia)

diasDoMes = Array(31, 28, 30, 30, 31, 30, 31, 30, 30, 31, 30, 31)

If dia = 31 Then
Verifica_dia = diasDoMes(mes - 1)
Else
Verifica_dia = dia
End If

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
