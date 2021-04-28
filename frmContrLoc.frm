VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmContrLoc 
   BorderStyle     =   0  'None
   Caption         =   "Super Imob - Cadastro de Contrato"
   ClientHeight    =   6585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data dtaContrato 
      Caption         =   "Contrato"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programa Imobiliária\Dados\Bdimobiliaria.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Contrato"
      Top             =   840
      Width           =   2895
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   6480
      TabIndex        =   42
      Top             =   480
      Width           =   2775
      Begin VB.CommandButton Command1 
         Caption         =   "<<< &Adicionar Clientes >>>"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Código Gerado:"
      Enabled         =   0   'False
      Height          =   735
      Left            =   1320
      TabIndex        =   30
      Top             =   480
      Width           =   2055
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   840
         TabIndex        =   31
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
   Begin VB.CommandButton Command3 
      Enabled         =   0   'False
      Height          =   615
      Left            =   240
      Picture         =   "frmContrLoc.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Novo Contrato"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Enabled         =   0   'False
      Height          =   615
      Left            =   1800
      Picture         =   "frmContrLoc.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Alterar Contrato"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Enabled         =   0   'False
      Height          =   615
      Left            =   3360
      Picture         =   "frmContrLoc.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Excluir Contrato"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Enabled         =   0   'False
      Height          =   615
      Left            =   4920
      Picture         =   "frmContrLoc.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Gravar Contrato"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Height          =   615
      Left            =   8040
      Picture         =   "frmContrLoc.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Sair"
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Frame Frame9 
      Caption         =   "Número:"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1095
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         DataField       =   "ID"
         DataSource      =   "dtaContrato"
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
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
   End
   Begin TabDlg.SSTab ssPainel 
      Height          =   4215
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1320
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Locador(a)"
      TabPicture(0)   =   "frmContrLoc.frx":154A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(1)=   "Frame10(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "1-Locatário(a)"
      TabPicture(1)   =   "frmContrLoc.frx":1566
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).Control(1)=   "Frame10(1)"
      Tab(1).Control(2)=   "Option1(2)"
      Tab(1).Control(3)=   "Option1(3)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "2-Locatário(a)"
      TabPicture(2)   =   "frmContrLoc.frx":1582
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(2)"
      Tab(2).Control(1)=   "Frame10(2)"
      Tab(2).Control(2)=   "Option1(4)"
      Tab(2).Control(3)=   "Option1(5)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "2-Fiador(a)"
      TabPicture(3)   =   "frmContrLoc.frx":159E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1(3)"
      Tab(3).Control(1)=   "Frame10(3)"
      Tab(3).Control(2)=   "Option1(6)"
      Tab(3).Control(3)=   "Option1(7)"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "1-Fiador(a)"
      TabPicture(4)   =   "frmContrLoc.frx":15BA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame10(4)"
      Tab(4).Control(1)=   "Frame1(4)"
      Tab(4).Control(2)=   "Option1(8)"
      Tab(4).Control(3)=   "Option1(9)"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Imóvel"
      TabPicture(5)   =   "frmContrLoc.frx":15D6
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame2"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Buscar"
      TabPicture(6)   =   "frmContrLoc.frx":15F2
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Frame3"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Frame6"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).ControlCount=   2
      Begin VB.Frame Frame6 
         Height          =   2775
         Left            =   240
         TabIndex        =   224
         Top             =   1200
         Width           =   8655
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "frmContrLoc.frx":160E
            Height          =   2415
            Left            =   120
            OleObjectBlob   =   "frmContrLoc.frx":1628
            TabIndex        =   225
            Top             =   240
            Width           =   8415
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Digite para buscar:"
         Height          =   615
         Left            =   240
         TabIndex        =   219
         Top             =   480
         Width           =   8655
         Begin VB.CommandButton Command2 
            Caption         =   ">>> Busc&ar >>>"
            Height          =   300
            Left            =   3120
            TabIndex        =   226
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Locatário"
            Height          =   195
            Left            =   7440
            TabIndex        =   223
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Locador"
            Height          =   195
            Left            =   6120
            TabIndex        =   222
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Código"
            Height          =   195
            Left            =   4920
            TabIndex        =   221
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text95 
            BackColor       =   &H00C0FFC0&
            Height          =   315
            Left            =   120
            TabIndex        =   220
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados do Imóvel:"
         Height          =   3495
         Left            =   -74760
         TabIndex        =   114
         Top             =   480
         Width           =   8655
         Begin VB.CommandButton Command7 
            BackColor       =   &H8000000B&
            Caption         =   "B&uscar imóvel cadastrado"
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
            Height          =   315
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   227
            Top             =   1800
            Width           =   3015
         End
         Begin VB.TextBox Text94 
            DataField       =   "Iptu"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   7440
            TabIndex        =   218
            Top             =   1800
            Width           =   1095
         End
         Begin VB.TextBox Text88 
            DataField       =   "Obs"
            DataSource      =   "dtaContrato"
            Height          =   1215
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   210
            Top             =   2160
            Width           =   8415
         End
         Begin VB.TextBox Text87 
            DataField       =   "Final"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6840
            TabIndex        =   209
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox Text86 
            DataField       =   "Inicio"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   3840
            TabIndex        =   208
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox Text85 
            DataField       =   "Prazo"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1560
            TabIndex        =   207
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox Text84 
            DataField       =   "Periodo"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   7200
            TabIndex        =   206
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox Text83 
            DataField       =   "Indice"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   3720
            TabIndex        =   205
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox Text82 
            DataField       =   "Dia"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   204
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox Text81 
            DataField       =   "Aluguel"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6840
            TabIndex        =   203
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox Text80 
            DataField       =   "BImovel"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1680
            TabIndex        =   202
            Top             =   720
            Width           =   4095
         End
         Begin VB.TextBox Text79 
            DataField       =   "Finalidade"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6360
            TabIndex        =   201
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox Text78 
            DataField       =   "ImovelLocado"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   960
            TabIndex        =   200
            Top             =   360
            Width           =   4215
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Iptu:"
            Height          =   195
            Left            =   7080
            TabIndex        =   217
            Top             =   1800
            Width           =   315
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Observação:"
            Height          =   195
            Left            =   120
            TabIndex        =   125
            Top             =   1920
            Width           =   915
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Término Contr.:"
            Height          =   195
            Left            =   5640
            TabIndex        =   124
            Top             =   1440
            Width           =   1080
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Início Contr.:"
            Height          =   195
            Left            =   2760
            TabIndex        =   123
            Top             =   1440
            Width           =   915
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Prazo do Contrato:"
            Height          =   195
            Left            =   120
            TabIndex        =   122
            Top             =   1440
            Width           =   1320
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Período p/ Atualização:"
            Height          =   195
            Left            =   5400
            TabIndex        =   121
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Índice de Atualização:"
            Height          =   195
            Left            =   2040
            TabIndex        =   120
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Vencimento:"
            Height          =   195
            Left            =   120
            TabIndex        =   119
            Top             =   1080
            Width           =   885
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Aluguél:"
            Height          =   195
            Left            =   6120
            TabIndex        =   118
            Top             =   720
            Width           =   570
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Finalidade:"
            Height          =   195
            Left            =   5520
            TabIndex        =   117
            Top             =   360
            Width           =   765
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Bairro / Cidade / UF:"
            Height          =   195
            Left            =   120
            TabIndex        =   116
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Left            =   120
            TabIndex        =   115
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cnpj:"
         Height          =   195
         Index           =   9
         Left            =   -69000
         TabIndex        =   10
         Top             =   1560
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cpf:"
         Height          =   195
         Index           =   8
         Left            =   -69720
         TabIndex        =   9
         Top             =   1560
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cnpj:"
         Height          =   195
         Index           =   7
         Left            =   -69000
         TabIndex        =   8
         Top             =   1560
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cpf:"
         Height          =   195
         Index           =   6
         Left            =   -69720
         TabIndex        =   7
         Top             =   1560
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cnpj:"
         Height          =   195
         Index           =   5
         Left            =   -69000
         TabIndex        =   6
         Top             =   1560
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cpf:"
         Height          =   195
         Index           =   4
         Left            =   -69720
         TabIndex        =   5
         Top             =   1560
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cnpj:"
         Height          =   195
         Index           =   3
         Left            =   -69000
         TabIndex        =   4
         Top             =   1560
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cpf:"
         Height          =   195
         Index           =   2
         Left            =   -69720
         TabIndex        =   3
         Top             =   1560
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados Fiador:"
         Enabled         =   0   'False
         Height          =   2295
         Index           =   4
         Left            =   -74760
         TabIndex        =   102
         Top             =   480
         Width           =   8655
         Begin VB.TextBox Text93 
            DataField       =   "afestcivil"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   216
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox Text72 
            DataField       =   "afcep"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   194
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox Text71 
            DataField       =   "afuf"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   4440
            TabIndex        =   193
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox Text70 
            DataField       =   "afcid"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   192
            Top             =   1800
            Width           =   2775
         End
         Begin VB.TextBox Text69 
            DataField       =   "afbairro"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6000
            TabIndex        =   191
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox Text68 
            DataField       =   "afend"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   190
            Top             =   1440
            Width           =   4215
         End
         Begin VB.TextBox Text67 
            DataField       =   "afcpf"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   189
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox Text66 
            DataField       =   "afrg"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   188
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox Text65 
            DataField       =   "afprof"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   187
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox Text64 
            DataField       =   "afnac"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   186
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text63 
            DataField       =   "afiador"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   720
            TabIndex        =   185
            Top             =   360
            Width           =   4575
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Uf:"
            Height          =   195
            Index           =   4
            Left            =   4080
            TabIndex        =   112
            Top             =   1800
            Width           =   210
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Cep:"
            Height          =   195
            Index           =   4
            Left            =   6120
            TabIndex        =   111
            Top             =   1800
            Width           =   330
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   110
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Index           =   4
            Left            =   5400
            TabIndex        =   109
            Top             =   1440
            Width           =   450
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   108
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rg:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   107
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil:"
            Height          =   195
            Index           =   4
            Left            =   5400
            TabIndex        =   106
            Top             =   720
            Width           =   870
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   105
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Index           =   4
            Left            =   5400
            TabIndex        =   104
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   103
            Top             =   360
            Width           =   465
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Conjuge:"
         Height          =   1095
         Index           =   4
         Left            =   -74760
         TabIndex        =   96
         Top             =   2880
         Width           =   8655
         Begin VB.TextBox Text77 
            DataField       =   "afconjcpf"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6720
            TabIndex        =   199
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox Text76 
            DataField       =   "afconjnac"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6720
            TabIndex        =   198
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox Text75 
            DataField       =   "afconjrg"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   3960
            TabIndex        =   197
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox Text74 
            DataField       =   "afconjprof"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   960
            TabIndex        =   196
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox Text73 
            DataField       =   "afconjuge"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   960
            TabIndex        =   195
            Top             =   360
            Width           =   4335
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   101
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Index           =   4
            Left            =   5520
            TabIndex        =   100
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   99
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Cpf:"
            Height          =   195
            Index           =   4
            Left            =   6240
            TabIndex        =   98
            Top             =   720
            Width           =   285
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Rg:"
            Height          =   195
            Index           =   4
            Left            =   3600
            TabIndex        =   97
            Top             =   720
            Width           =   255
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Conjuge:"
         Height          =   1095
         Index           =   3
         Left            =   -74760
         TabIndex        =   90
         Top             =   2880
         Width           =   8655
         Begin VB.TextBox Text62 
            DataField       =   "fconjcpf"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6720
            TabIndex        =   184
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox Text61 
            DataField       =   "fconjrg"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   3960
            TabIndex        =   183
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox Text60 
            DataField       =   "fconjprof"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   960
            TabIndex        =   182
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox Text59 
            DataField       =   "fconjnac"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6720
            TabIndex        =   181
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox Text58 
            DataField       =   "fconjuge"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   960
            TabIndex        =   180
            Top             =   360
            Width           =   4335
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   95
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Index           =   3
            Left            =   5520
            TabIndex        =   94
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   93
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Cpf:"
            Height          =   195
            Index           =   3
            Left            =   6240
            TabIndex        =   92
            Top             =   720
            Width           =   285
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Rg:"
            Height          =   195
            Index           =   3
            Left            =   3600
            TabIndex        =   91
            Top             =   720
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados Fiador:"
         Enabled         =   0   'False
         Height          =   2295
         Index           =   3
         Left            =   -74760
         TabIndex        =   79
         Top             =   480
         Width           =   8655
         Begin VB.TextBox Text92 
            DataField       =   "festcivil"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   215
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox Text57 
            DataField       =   "fcep"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   179
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox Text56 
            DataField       =   "fuf"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   4440
            TabIndex        =   178
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox Text55 
            DataField       =   "fcid"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   177
            Top             =   1800
            Width           =   2775
         End
         Begin VB.TextBox Text54 
            DataField       =   "fbairro"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6000
            TabIndex        =   176
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox Text53 
            DataField       =   "fend"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   175
            Top             =   1440
            Width           =   4215
         End
         Begin VB.TextBox Text52 
            DataField       =   "fcpf"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   174
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox Text51 
            DataField       =   "frg"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   173
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox Text50 
            DataField       =   "fprof"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   172
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox Text49 
            DataField       =   "fnac"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   171
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text48 
            DataField       =   "fiador"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   720
            TabIndex        =   170
            Top             =   360
            Width           =   4575
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   89
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Index           =   3
            Left            =   5400
            TabIndex        =   88
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   87
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil:"
            Height          =   195
            Index           =   3
            Left            =   5400
            TabIndex        =   86
            Top             =   720
            Width           =   870
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rg:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   85
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   84
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Index           =   3
            Left            =   5400
            TabIndex        =   83
            Top             =   1440
            Width           =   450
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   82
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Cep:"
            Height          =   195
            Index           =   3
            Left            =   6120
            TabIndex        =   81
            Top             =   1800
            Width           =   330
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Uf:"
            Height          =   195
            Index           =   3
            Left            =   4080
            TabIndex        =   80
            Top             =   1800
            Width           =   210
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Conjuge:"
         Height          =   1095
         Index           =   2
         Left            =   -74760
         TabIndex        =   73
         Top             =   2880
         Width           =   8655
         Begin VB.TextBox Text47 
            DataField       =   "blocatarioconjcpf"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6720
            TabIndex        =   169
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox Text46 
            DataField       =   "blocatarioconjrg"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   3960
            TabIndex        =   168
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox Text45 
            DataField       =   "blocatarioconjprof"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   960
            TabIndex        =   167
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox Text44 
            DataField       =   "blocatarioconjnac"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6720
            TabIndex        =   166
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox Text43 
            DataField       =   "blocatarioconjuge"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   960
            TabIndex        =   165
            Top             =   360
            Width           =   4335
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Rg:"
            Height          =   195
            Index           =   2
            Left            =   3600
            TabIndex        =   78
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Cpf:"
            Height          =   195
            Index           =   2
            Left            =   6240
            TabIndex        =   77
            Top             =   720
            Width           =   285
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   76
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Index           =   2
            Left            =   5520
            TabIndex        =   75
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   74
            Top             =   360
            Width           =   465
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados Locatário:"
         Enabled         =   0   'False
         Height          =   2295
         Index           =   2
         Left            =   -74760
         TabIndex        =   62
         Top             =   480
         Width           =   8655
         Begin VB.TextBox Text91 
            DataField       =   "blocatarioestcivil"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   214
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox Text35 
            DataField       =   "blocatarioprof"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   211
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox Text42 
            DataField       =   "blocatariocep"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   164
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox Text41 
            DataField       =   "blocatariouf"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   4440
            TabIndex        =   163
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox Text40 
            DataField       =   "blocatariocid"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   162
            Top             =   1800
            Width           =   2775
         End
         Begin VB.TextBox Text39 
            DataField       =   "blocatariobairro"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6000
            TabIndex        =   161
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox Text38 
            DataField       =   "blocatarioend"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   160
            Top             =   1440
            Width           =   4215
         End
         Begin VB.TextBox Text37 
            DataField       =   "blocatariocpf"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   159
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox Text36 
            DataField       =   "blocatariorg"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   158
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox Text34 
            DataField       =   "blocatarionac"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   157
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text33 
            DataField       =   "blocatario"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   720
            TabIndex        =   156
            Top             =   360
            Width           =   4575
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Uf:"
            Height          =   195
            Index           =   2
            Left            =   4080
            TabIndex        =   72
            Top             =   1800
            Width           =   210
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Cep:"
            Height          =   195
            Index           =   2
            Left            =   6120
            TabIndex        =   71
            Top             =   1800
            Width           =   330
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   70
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Index           =   2
            Left            =   5400
            TabIndex        =   69
            Top             =   1440
            Width           =   450
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   68
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rg:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   67
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil:"
            Height          =   195
            Index           =   2
            Left            =   5400
            TabIndex        =   66
            Top             =   720
            Width           =   870
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   65
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Index           =   2
            Left            =   5400
            TabIndex        =   64
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   63
            Top             =   360
            Width           =   465
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Conjuge:"
         Height          =   1095
         Index           =   1
         Left            =   -74760
         TabIndex        =   56
         Top             =   2880
         Width           =   8655
         Begin VB.TextBox Text32 
            DataField       =   "alocatarioconjcpf"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6720
            TabIndex        =   155
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox Text31 
            DataField       =   "alocatarioconjrg"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   3960
            TabIndex        =   154
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox Text30 
            DataField       =   "alocatarioconjprof"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   960
            TabIndex        =   153
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox Text29 
            DataField       =   "aLocatarioconjnac"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6720
            TabIndex        =   152
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox Text28 
            DataField       =   "alocatarioconjuge"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   960
            TabIndex        =   151
            Top             =   360
            Width           =   4335
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Index           =   1
            Left            =   5520
            TabIndex        =   60
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   59
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Cpf:"
            Height          =   195
            Index           =   1
            Left            =   6240
            TabIndex        =   58
            Top             =   720
            Width           =   285
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Rg:"
            Height          =   195
            Index           =   1
            Left            =   3600
            TabIndex        =   57
            Top             =   720
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados Locatário:"
         Enabled         =   0   'False
         Height          =   2295
         Index           =   1
         Left            =   -74760
         TabIndex        =   45
         Top             =   480
         Width           =   8655
         Begin VB.TextBox Text90 
            DataField       =   "aLocatarioestcivil"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   213
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox Text27 
            DataField       =   "alocatariocep"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   150
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox Text26 
            DataField       =   "alocatariouf"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   4440
            TabIndex        =   149
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox Text25 
            DataField       =   "alocatariocid"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   148
            Top             =   1800
            Width           =   2775
         End
         Begin VB.TextBox Text24 
            DataField       =   "aLocatariobairro"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6000
            TabIndex        =   147
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox Text23 
            DataField       =   "aLocatarioend"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   146
            Top             =   1440
            Width           =   4215
         End
         Begin VB.TextBox Text22 
            DataField       =   "aLocatariocpf"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   145
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox Text21 
            DataField       =   "aLocatariorg"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   144
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox Text20 
            DataField       =   "aLocatarioprof"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   143
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox Text19 
            DataField       =   "aLocatarionac"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   142
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text18 
            DataField       =   "aLocatario"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   720
            TabIndex        =   141
            Top             =   360
            Width           =   4575
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   55
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Index           =   1
            Left            =   5400
            TabIndex        =   54
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   53
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil:"
            Height          =   195
            Index           =   1
            Left            =   5400
            TabIndex        =   52
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
            TabIndex        =   51
            Top             =   1080
            Width           =   255
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
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Index           =   1
            Left            =   5400
            TabIndex        =   49
            Top             =   1440
            Width           =   450
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   48
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Cep:"
            Height          =   195
            Index           =   1
            Left            =   6120
            TabIndex        =   47
            Top             =   1800
            Width           =   330
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Uf:"
            Height          =   195
            Index           =   1
            Left            =   4080
            TabIndex        =   46
            Top             =   1800
            Width           =   210
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Conjuge:"
         Height          =   1095
         Index           =   0
         Left            =   -74760
         TabIndex        =   37
         Top             =   2880
         Width           =   8655
         Begin VB.TextBox Text17 
            DataField       =   "LocConjCpf"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6720
            TabIndex        =   140
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox Text16 
            DataField       =   "LocConjRg"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   3960
            TabIndex        =   139
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox Text15 
            DataField       =   "LocConjProf"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   960
            TabIndex        =   138
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox Text14 
            DataField       =   "LocConjNac"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6720
            TabIndex        =   137
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox Text13 
            DataField       =   "ConjugeLocador"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   960
            TabIndex        =   136
            Top             =   360
            Width           =   4335
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Rg:"
            Height          =   195
            Index           =   0
            Left            =   3600
            TabIndex        =   44
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Cpf:"
            Height          =   195
            Index           =   0
            Left            =   6240
            TabIndex        =   41
            Top             =   720
            Width           =   285
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Index           =   0
            Left            =   5520
            TabIndex        =   39
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   465
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados Locador:"
         Enabled         =   0   'False
         Height          =   2295
         Index           =   0
         Left            =   -74760
         TabIndex        =   14
         Top             =   480
         Width           =   8655
         Begin VB.TextBox Text89 
            DataField       =   "ELocador"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   212
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox Text12 
            DataField       =   "Ceplocador"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   135
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox Text11 
            DataField       =   "uflocador"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   4440
            TabIndex        =   134
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox Text10 
            DataField       =   "Cidlocador"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   133
            Top             =   1800
            Width           =   2775
         End
         Begin VB.TextBox Text9 
            DataField       =   "BLocador"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6000
            TabIndex        =   132
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox Text8 
            DataField       =   "EndLocador"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   131
            Top             =   1440
            Width           =   4215
         End
         Begin VB.TextBox Text7 
            DataField       =   "CpfLocador"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   130
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox Text6 
            DataField       =   "RgLocador"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   129
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox Text5 
            DataField       =   "ProfLocador"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   1080
            TabIndex        =   128
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox Text4 
            DataField       =   "NLocador"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   6600
            TabIndex        =   127
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text3 
            DataField       =   "Locador"
            DataSource      =   "dtaContrato"
            Height          =   285
            Left            =   720
            TabIndex        =   126
            Top             =   360
            Width           =   4575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Cnpj:"
            Height          =   195
            Index           =   1
            Left            =   5760
            TabIndex        =   2
            Top             =   1080
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Cpf:"
            Height          =   195
            Index           =   0
            Left            =   5040
            TabIndex        =   1
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Uf:"
            Height          =   195
            Index           =   0
            Left            =   4080
            TabIndex        =   24
            Top             =   1800
            Width           =   210
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Cep:"
            Height          =   195
            Index           =   0
            Left            =   6120
            TabIndex        =   23
            Top             =   1800
            Width           =   330
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Index           =   0
            Left            =   5400
            TabIndex        =   21
            Top             =   1440
            Width           =   450
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   20
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
            TabIndex        =   19
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil:"
            Height          =   195
            Index           =   0
            Left            =   5400
            TabIndex        =   18
            Top             =   720
            Width           =   870
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Profissão:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidade:"
            Height          =   195
            Index           =   0
            Left            =   5400
            TabIndex        =   16
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   465
         End
      End
   End
   Begin VB.Frame Frame8 
      Height          =   975
      Left            =   120
      TabIndex        =   33
      Top             =   5520
      Width           =   9135
      Begin VB.CommandButton Command9 
         Enabled         =   0   'False
         Height          =   615
         Left            =   6360
         Picture         =   "frmContrLoc.frx":2363
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Buscar Contrato"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Label14"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   4080
      TabIndex        =   113
      Top             =   480
      Width           =   750
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   4440
      TabIndex        =   43
      Top             =   120
      Width           =   480
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
      Caption         =   "CONTRATO DE LOCAÇÃO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   210
      Left            =   120
      TabIndex        =   36
      Top             =   120
      Width           =   2505
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
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   7320
      TabIndex        =   35
      Top             =   120
      Width           =   885
   End
End
Attribute VB_Name = "frmContrLoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ws As Workspace
Dim Dados As DAO.Database
Dim Tabela As DAO.Recordset
Public Sql As String

Private Sub Command10_Click()
ssPainel.Tab = 0
End Sub

Private Sub Command2_Click()
On Error Resume Next

If Option2.Value = True Then
    If Text95.Text = "" Then
        dtaContrato.RecordSource = "SELECT * FROM CONTRATO"
        dtaContrato.Refresh
        Exit Sub
    End If

    dtaContrato.RecordSource = "SELECT * FROM CONTRATO WHERE ID Like '" & Text95.Text & "*'"
    dtaContrato.Refresh

ElseIf Option3.Value = True Then
    If Text95.Text = "" Then
        dtaContrato.RecordSource = "SELECT * FROM CONTRATO"
        dtaContrato.Refresh
        Exit Sub
    End If

    dtaContrato.RecordSource = "SELECT * FROM CONTRATO WHERE LOCADOR Like '" & Text95.Text & "*'"
    dtaContrato.Refresh
    
ElseIf Option4.Value = True Then
    If Text95.Text = "" Then
        dtaContrato.RecordSource = "SELECT * FROM CONTRATO"
        dtaContrato.Refresh
        Exit Sub
    End If

    dtaContrato.RecordSource = "SELECT * FROM CONTRATO WHERE ALOCATARIO Like '" & Text95.Text & "*'"
    dtaContrato.Refresh
End If
End Sub

Private Sub Command1_Click()
frmClientes.Command1.Enabled = True
frmClientes.Show 1
End Sub

Private Sub Command3_Click()
Dim Novo As String

Set Dados = OpenDatabase(App.Path & "\Dados\Bdimobiliaria.MDB")
Set Tabela = Dados.OpenRecordset("Contrato", dbOpenTable)

Tabela.Index = "ID"
Tabela.MoveLast
Novo = Tabela!ID
Tabela.Seek "=", Novo
If Tabela.NoMatch = False Then
    Novo = Novo + 1
End If
    Dados.Close
    
    dtaContrato.Recordset.AddNew
    CorTxtAberto
    AbreTab0
    AbreTab1
    AbreTab2
    AbreTab3
    AbreTab4
    AbreTab5
    LimpaCaixas
    Text1 = Novo
    Text3.SetFocus
    ssPainel.Tab = 0
    Frame1(0).Enabled = True
    Frame2.Enabled = True
    Command8.Picture = LoadPicture(App.Path & "\Ícones\TRFFC14.ico")
    Command8.ToolTipText = "Cancelar Novo"
    Command1.Enabled = True
    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = True
    Command9.Enabled = False
    Text83.Text = "IGPM/FGV"
    Text84.Text = "ANUAL"
    
End Sub

Private Sub Command4_Click()

dtaContrato.Recordset.Edit
ssPainel.Tab = 0
Frame1(0).Enabled = True
CorTxtAberto
AbreTab0
AbreTab1
AbreTab2
AbreTab3
AbreTab4
AbreTab5
Command8.Picture = LoadPicture(App.Path & "\Ícones\TRFFC14.ico")
Command8.ToolTipText = "Cancelar Novo"
Text3.SetFocus
Command1.Enabled = True
Command6.Enabled = True
Command4.Enabled = False
Command5.Enabled = False
Command9.Enabled = False

End Sub

Private Sub Command5_Click()
 If MsgBox("Confirma Exclusão do Contrato  -> " & dtaContrato.Recordset![ID], vbQuestion + vbYesNo, "Excluir Contrato") = vbYes Then
      dtaContrato.Recordset.Delete
      dtaContrato.Refresh
    End If
End Sub

Private Sub Command6_Click()

    ConverteTexto
    dtaContrato.UpdateRecord
    dtaContrato.Recordset.Bookmark = dtaContrato.Recordset.LastModified
    dtaContrato.Refresh
    MsgBox ("Dados cadastrados com sucesso!")
    FechaTab0
    FechaTab1
    FechaTab2
    FechaTab3
    FechaTab4
    FechaTab5
    Command1.Caption = "<<< &Adicionar Clientes >>>"
    Command6.Enabled = False
    Command8.ToolTipText = "Sair"
    Command8.Picture = LoadPicture(App.Path & "\Ícones\ARW10NE.ico")
    Command3.Enabled = True
    CorTxtFechado
    Frame4.Enabled = False
    Command4.Enabled = True
    Command5.Enabled = True
    Command9.Enabled = True
    ssPainel.Tab = 0

End Sub

Private Sub Command7_Click()
frmImoveis.cmdNovo.Enabled = False
frmImoveis.cmdAlterar.Enabled = False
frmImoveis.cmdGravar.Enabled = False
frmImoveis.cmdRemover.Enabled = False
frmImoveis.cmdImovel.Enabled = True
frmImoveis.Show 1
End Sub

Private Sub Command8_Click()

If Command8.ToolTipText = "Cancelar Novo" Then
    If dtaContrato.Recordset.RecordCount = 0 Then
        dtaContrato.Recordset.CancelUpdate
        Command8.ToolTipText = "Sair"
        Command8.Picture = LoadPicture(App.Path & "\Ícones\ARW10NE.ico")
        CorTxtFechado
        FechaTab0
        FechaTab1
        FechaTab2
        FechaTab3
        FechaTab4
        FechaTab5
        ssPainel.Tab = 0
        Text3.Enabled = False
        Command1.Enabled = False
        Command1.Caption = "<<< &Adicionar Clientes >>>"
        Command3.Enabled = True
        Command4.Enabled = True
        Command5.Enabled = True
        Command6.Enabled = False
        Command9.Enabled = True
        
    Else
        
        dtaContrato.Recordset.CancelUpdate
        dtaContrato.Recordset.MoveLast
        Command8.ToolTipText = "Sair"
        Command8.Picture = LoadPicture(App.Path & "\Ícones\ARW10NE.ico")
        CorTxtFechado
        FechaTab0
        FechaTab1
        FechaTab2
        FechaTab3
        FechaTab4
        FechaTab5
        ssPainel.Tab = 0
        Text3.Enabled = False
        Command1.Enabled = False
        Command1.Caption = "<<< &Adicionar Clientes >>>"
        Command3.Enabled = True
        Command4.Enabled = True
        Command5.Enabled = True
        Command6.Enabled = False
        Command9.Enabled = True
    End If
Else
    
    If MsgBox("Quer sair do Cadastro de Contrato?", vbYesNo, "Sair do Cadastro") = vbYes Then
    Unload frmContrLoc
    RedefineFormPrincipal
  Else
    Exit Sub
End If
End If
End Sub

Private Sub Command9_Click()
ssPainel.Tab = 6
End Sub

Private Sub DBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
On Error GoTo erro

If ColIndex >= 0 And ColIndex <= 4 Then
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
If KeyAscii = vbKeyReturn Then
    SendKeys ("{TAB}")
    KeyAscii = 0
End If
End Sub

Private Sub Form_Load()

    dtaContrato.DatabaseName = App.Path & "\Dados\Bdimobiliaria.MDB"
    dtaContrato.RecordSource = "Contrato"
    dtaContrato.Refresh
    Label8.Caption = "Hoje é " & Date
    Label14.Caption = "Dados do Locador"

    Frame2.Enabled = False
    Frame1(0).Enabled = False
    Frame1(1).Enabled = False
    Frame1(2).Enabled = False
    Frame1(3).Enabled = False
    Frame1(4).Enabled = False
    Frame10(0).Enabled = False
    Frame10(1).Enabled = False
    Frame10(2).Enabled = False
    Frame10(3).Enabled = False
    Frame10(4).Enabled = False

    FechaTab0
    FechaTab1
    FechaTab2
    FechaTab3
    FechaTab4
    FechaTab5

    CorTxtFechado
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command8.Enabled = True
    Command9.Enabled = True

    ssPainel.Tab = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmFundo.Enabled = True
End Sub

Private Sub Option2_Click()
Text95 = ""
Text95.SetFocus
dtaContrato.Refresh
End Sub

Private Sub Option3_Click()
Text95 = ""
Text95.SetFocus
dtaContrato.Refresh
End Sub

Private Sub Option4_Click()
Text95 = ""
Text95.SetFocus
dtaContrato.Refresh
End Sub

Private Sub ssPainel_Click(PreviousTab As Integer)

If ssPainel.Tab = 0 Then
    Label14.Caption = "Dados do Locador"
End If
If ssPainel.Tab = 1 Then
    Label14.Caption = "Dados do Locatário"
End If
If ssPainel.Tab = 2 Then
    Label14.Caption = "Dados dos Locatários"
End If
If ssPainel.Tab = 3 Then
    Label14.Caption = "Dados do Fiador"
End If
If ssPainel.Tab = 4 Then
    Label14.Caption = "Dados dos Fiadores"
End If
If ssPainel.Tab = 5 Then
    Label14.Caption = "Dados do Imóvel"
End If
End Sub

Private Function CorTxtFechado()
Dim TxtBox As Object

For Each TxtBox In Me.Controls
    If TypeOf TxtBox Is TextBox Then
    TxtBox.BackColor = &H8000000B
End If
Next TxtBox
Text1.BackColor = &HFFFF&
Text2.BackColor = &HFFFF&
Text95.BackColor = &HC0FFC0
End Function

Private Function CorTxtAberto()
Dim TxtBox As Object
Dim Combos As Object

For Each TxtBox In Me.Controls
    If TypeOf TxtBox Is TextBox Then
    TxtBox.BackColor = &H80000005
End If
Next TxtBox
Text1.BackColor = &HFFFF&
Text2.BackColor = &HFFFF&

End Function

Private Function AbreTab0()

Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
Text14.Enabled = True
Text15.Enabled = True
Text16.Enabled = True
Text17.Enabled = True
Text89.Enabled = True
Option1(0).Enabled = True
Option1(0).Value = False
Option1(1).Enabled = True
Option1(1).Value = False
Frame1(0).Enabled = True
Frame10(0).Enabled = True

End Function

Private Function FechaTab0()

Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
Text16.Enabled = False
Text17.Enabled = False
Text89.Enabled = False
Option1(0).Enabled = False
Option1(0).Value = False
Option1(1).Enabled = False
Option1(1).Value = False
Frame1(0).Enabled = False
Frame10(0).Enabled = False

End Function

Private Function AbreTab1()

Text18.Enabled = True
Text19.Enabled = True
Text20.Enabled = True
Text21.Enabled = True
Text22.Enabled = True
Text23.Enabled = True
Text24.Enabled = True
Text25.Enabled = True
Text26.Enabled = True
Text27.Enabled = True
Text28.Enabled = True
Text29.Enabled = True
Text30.Enabled = True
Text31.Enabled = True
Text32.Enabled = True
Text90.Enabled = True
Option1(2).Enabled = False
Option1(2).Value = False
Option1(3).Enabled = False
Option1(3).Value = False
Frame1(1).Enabled = True
Frame10(1).Enabled = True

End Function

Private Function FechaTab1()

Text18.Enabled = False
Text19.Enabled = False
Text20.Enabled = False
Text21.Enabled = False
Text22.Enabled = False
Text23.Enabled = False
Text24.Enabled = False
Text25.Enabled = False
Text26.Enabled = False
Text27.Enabled = False
Text28.Enabled = False
Text29.Enabled = False
Text30.Enabled = False
Text31.Enabled = False
Text32.Enabled = False
Text90.Enabled = False
Option1(2).Enabled = False
Option1(2).Value = False
Option1(3).Enabled = False
Option1(3).Value = False
Frame1(1).Enabled = False
Frame10(1).Enabled = False

End Function

Private Function AbreTab2()

Text33.Enabled = True
Text34.Enabled = True
Text35.Enabled = True
Text36.Enabled = True
Text37.Enabled = True
Text38.Enabled = True
Text39.Enabled = True
Text40.Enabled = True
Text41.Enabled = True
Text42.Enabled = True
Text43.Enabled = True
Text44.Enabled = True
Text45.Enabled = True
Text46.Enabled = True
Text47.Enabled = True
Text91.Enabled = True
Option1(4).Enabled = True
Option1(4).Value = False
Option1(5).Enabled = True
Option1(5).Value = False
Frame1(2).Enabled = True
Frame10(2).Enabled = True

End Function

Private Function FechaTab2()

Text33.Enabled = False
Text34.Enabled = False
Text35.Enabled = False
Text36.Enabled = False
Text37.Enabled = False
Text38.Enabled = False
Text39.Enabled = False
Text40.Enabled = False
Text41.Enabled = False
Text42.Enabled = False
Text43.Enabled = False
Text44.Enabled = False
Text45.Enabled = False
Text46.Enabled = False
Text47.Enabled = False
Text91.Enabled = False
Option1(4).Enabled = False
Option1(4).Value = False
Option1(5).Enabled = False
Option1(5).Value = False
Frame1(2).Enabled = False
Frame10(2).Enabled = False

End Function

Private Function AbreTab3()

Text48.Enabled = True
Text49.Enabled = True
Text50.Enabled = True
Text51.Enabled = True
Text52.Enabled = True
Text53.Enabled = True
Text54.Enabled = True
Text55.Enabled = True
Text56.Enabled = True
Text57.Enabled = True
Text58.Enabled = True
Text59.Enabled = True
Text60.Enabled = True
Text61.Enabled = True
Text62.Enabled = True
Text92.Enabled = True
Option1(6).Enabled = True
Option1(6).Value = False
Option1(7).Enabled = True
Option1(7).Value = False
Frame1(3).Enabled = True
Frame10(3).Enabled = True

End Function

Private Function FechaTab3()

Text48.Enabled = False
Text49.Enabled = False
Text50.Enabled = False
Text51.Enabled = False
Text52.Enabled = False
Text53.Enabled = False
Text54.Enabled = False
Text55.Enabled = False
Text56.Enabled = False
Text57.Enabled = False
Text58.Enabled = False
Text59.Enabled = False
Text60.Enabled = False
Text61.Enabled = False
Text62.Enabled = False
Text92.Enabled = False
Option1(6).Enabled = False
Option1(6).Value = False
Option1(7).Enabled = False
Option1(7).Value = False
Frame1(3).Enabled = False
Frame10(3).Enabled = False

End Function

Private Function AbreTab4()

Text63.Enabled = True
Text64.Enabled = True
Text65.Enabled = True
Text66.Enabled = True
Text67.Enabled = True
Text68.Enabled = True
Text69.Enabled = True
Text70.Enabled = True
Text71.Enabled = True
Text72.Enabled = True
Text73.Enabled = True
Text74.Enabled = True
Text75.Enabled = True
Text76.Enabled = True
Text77.Enabled = True
Text93.Enabled = True
Option1(8).Enabled = True
Option1(8).Value = False
Option1(9).Enabled = True
Option1(9).Value = False
Frame1(4).Enabled = True
Frame10(4).Enabled = True

End Function

Private Function FechaTab4()

Text63.Enabled = False
Text64.Enabled = False
Text65.Enabled = False
Text66.Enabled = False
Text67.Enabled = False
Text68.Enabled = False
Text69.Enabled = False
Text70.Enabled = False
Text71.Enabled = False
Text72.Enabled = False
Text73.Enabled = False
Text74.Enabled = False
Text75.Enabled = False
Text76.Enabled = False
Text77.Enabled = False
Text93.Enabled = False
Option1(8).Enabled = False
Option1(8).Value = False
Option1(9).Enabled = False
Option1(9).Value = False
Frame1(4).Enabled = False
Frame10(4).Enabled = False

End Function

Private Function AbreTab5()

Text78.Enabled = True
Text79.Enabled = True
Text80.Enabled = True
Text81.Enabled = True
Text82.Enabled = True
Text83.Enabled = True
Text84.Enabled = True
Text85.Enabled = True
Text86.Enabled = True
Text87.Enabled = True
Text88.Enabled = True
Text94.Enabled = True
Frame2.Enabled = True
Command7.Enabled = True

End Function

Private Function FechaTab5()

Text78.Enabled = False
Text79.Enabled = False
Text80.Enabled = False
Text81.Enabled = False
Text82.Enabled = False
Text83.Enabled = False
Text84.Enabled = False
Text85.Enabled = False
Text86.Enabled = False
Text87.Enabled = False
Text88.Enabled = False
Text94.Enabled = False
Command7.Enabled = False

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
Text2 = "KC" & Format(Text1, "000")
End Sub


Private Function Tab0()

Text3 = dtaClientes.Recordset("Nome")
Text4 = dtaClientes.Recordset("Nacionalidade")
Text5 = dtaClientes.Recordset("Profissao")
Text89 = dtaClientes.Recordset("EstadoCivil")
Text6 = dtaClientes.Recordset("Rg")
Text7 = dtaClientes.Recordset("Cpf")
Text8 = dtaClientes.Recordset("Ruanoti")
Text9 = dtaClientes.Recordset("Bairronoti")
Text10 = dtaClientes.Recordset("CidadeNoti")
Text11 = dtaClientes.Recordset("Estado")
Text12 = dtaClientes.Recordset("Cepnoti")
Text13 = dtaClientes.Recordset("Conjuge")
Text14 = dtaClientes.Recordset("Nacionalconjuge")
Text15 = dtaClientes.Recordset("ProfConjuge")
Text16 = dtaClientes.Recordset("Rgconjuge")
Text17 = dtaClientes.Recordset("Cpfconjuge")

End Function


Private Function Tab1()

Text18 = dtaClientes.Recordset("Nome")
Text19 = dtaClientes.Recordset("Nacionalidade")
Text20 = dtaClientes.Recordset("Profissao")
Text90 = dtaClientes.Recordset("EstadoCivil")
Text21 = dtaClientes.Recordset("Rg")
Text22 = dtaClientes.Recordset("Cpf")
Text23 = dtaClientes.Recordset("RuaNoti")
Text24 = dtaClientes.Recordset("BairroNoti")
Text25 = dtaClientes.Recordset("CidadeNoti")
Text26 = dtaClientes.Recordset("Estado")
Text27 = dtaClientes.Recordset("Cepnoti")
Text28 = dtaClientes.Recordset("Conjuge")
Text29 = dtaClientes.Recordset("NacionalConjuge")
Text30 = dtaClientes.Recordset("ProfConjuge")
Text31 = dtaClientes.Recordset("RgConjuge")
Text32 = dtaClientes.Recordset("CpfConjuge")

End Function


Private Function Tab2()

Text33 = dtaClientes.Recordset("Nome")
Text34 = dtaClientes.Recordset("Nacionalidade")
Text35 = dtaClientes.Recordset("Profissao")
Text91 = dtaClientes.Recordset("EstadoCivil")
Text36 = dtaClientes.Recordset("Rg")
Text37 = dtaClientes.Recordset("Cpf")
Text38 = dtaClientes.Recordset("RuaNoti")
Text39 = dtaClientes.Recordset("BairroNoti")
Text40 = dtaClientes.Recordset("CidadeNoti")
Text41 = dtaClientes.Recordset("Estado")
Text42 = dtaClientes.Recordset("Cepnoti")
Text43 = dtaClientes.Recordset("Conjuge")
Text44 = dtaClientes.Recordset("NacionalConjuge")
Text45 = dtaClientes.Recordset("ProfConjuge")
Text46 = dtaClientes.Recordset("RgConjuge")
Text47 = dtaClientes.Recordset("CpfConjuge")

End Function

Private Function tab3()

Text48 = dtaClientes.Recordset("Nome")
Text49 = dtaClientes.Recordset("Nacionalidade")
Text50 = dtaClientes.Recordset("Profissao")
Text92 = dtaClientes.Recordset("EstadoCivil")
Text51 = dtaClientes.Recordset("Rg")
Text52 = dtaClientes.Recordset("Cpf")
Text53 = dtaClientes.Recordset("Ruanoti")
Text54 = dtaClientes.Recordset("BairroNoti")
Text55 = dtaClientes.Recordset("CidadeNoti")
Text56 = dtaClientes.Recordset("Estado")
Text57 = dtaClientes.Recordset("CepNoti")
Text58 = dtaClientes.Recordset("Conjuge")
Text59 = dtaClientes.Recordset("NacionalConjuge")
Text60 = dtaClientes.Recordset("ProfConjuge")
Text61 = dtaClientes.Recordset("RgConjuge")
Text62 = dtaClientes.Recordset("CpfConjuge")

End Function

Private Function Tab4()

Text63 = dtaClientes.Recordset("Nome")
Text64 = dtaClientes.Recordset("Nacionalidade")
Text65 = dtaClientes.Recordset("Profissao")
Text93 = dtaClientes.Recordset("EstadoCivil")
Text66 = dtaClientes.Recordset("Rg")
Text67 = dtaClientes.Recordset("Cpf")
Text68 = dtaClientes.Recordset("RuaNoti")
Text69 = dtaClientes.Recordset("BairroNoti")
Text70 = dtaClientes.Recordset("CidadeNoti")
Text71 = dtaClientes.Recordset("Estado")
Text72 = dtaClientes.Recordset("CepNoti")
Text73 = dtaClientes.Recordset("Conjuge")
Text74 = dtaClientes.Recordset("NacionalConjuge")
Text75 = dtaClientes.Recordset("ProfConjuge")
Text76 = dtaClientes.Recordset("Rgconjuge")
Text77 = dtaClientes.Recordset("Cpfconjuge")
End Function

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text11.SetFocus
  End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text12.SetFocus
  End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text13.SetFocus
  End If
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text14.SetFocus
  End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text15.SetFocus
  End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text16.SetFocus
  End If
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text17.SetFocus
  End If
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ssPainel.Tab = 1
     Text18.SetFocus
  End If
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text19.SetFocus
  End If
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text20.SetFocus
  End If
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text90.SetFocus
  End If
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text22.SetFocus
  End If
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text23.SetFocus
  End If
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text24.SetFocus
  End If
End Sub

Private Sub Text24_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text25.SetFocus
  End If
End Sub

Private Sub Text25_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text26.SetFocus
  End If
End Sub

Private Sub Text26_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text27.SetFocus
  End If
End Sub

Private Sub Text27_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text28.SetFocus
  End If
End Sub

Private Sub Text28_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text29.SetFocus
  End If
End Sub

Private Sub Text29_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text30.SetFocus
  End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text4.SetFocus
  End If
End Sub

Private Sub Text30_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text31.SetFocus
  End If
End Sub

Private Sub Text31_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text32.SetFocus
  End If
End Sub

Private Sub Text32_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     ssPainel.Tab = 2
     Text33.SetFocus
  End If
End Sub

Private Sub Text33_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text34.SetFocus
  End If
End Sub

Private Sub Text34_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text35.SetFocus
  End If
End Sub

Private Sub Text35_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text91.SetFocus
  End If
End Sub

Private Sub Text36_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text37.SetFocus
  End If
End Sub

Private Sub Text37_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text38.SetFocus
  End If
End Sub

Private Sub Text38_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text39.SetFocus
  End If
End Sub

Private Sub Text39_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text40.SetFocus
  End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text5.SetFocus
  End If
End Sub

Private Sub Text40_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text41.SetFocus
  End If
End Sub

Private Sub Text41_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text42.SetFocus
  End If
End Sub

Private Sub Text42_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text43.SetFocus
  End If
End Sub

Private Sub Text43_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text44.SetFocus
  End If
End Sub

Private Sub Text44_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text45.SetFocus
  End If
End Sub

Private Sub Text45_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text46.SetFocus
  End If
End Sub

Private Sub Text46_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text47.SetFocus
  End If
End Sub

Private Sub Text47_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     ssPainel.Tab = 3
     Text48.SetFocus
  End If
End Sub

Private Sub Text48_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text49.SetFocus
  End If
End Sub

Private Sub Text49_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text50.SetFocus
  End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text89.SetFocus
  End If
End Sub

Private Sub Text50_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text92.SetFocus
  End If
End Sub

Private Sub Text51_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text52.SetFocus
  End If
End Sub

Private Sub Text52_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text53.SetFocus
  End If
End Sub

Private Sub Text53_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text54.SetFocus
  End If
End Sub

Private Sub Text54_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text55.SetFocus
  End If
End Sub

Private Sub Text55_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text56.SetFocus
  End If
End Sub

Private Sub Text56_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text57.SetFocus
  End If
End Sub

Private Sub Text57_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text58.SetFocus
  End If
End Sub

Private Sub Text58_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text59.SetFocus
  End If
End Sub

Private Sub Text59_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text60.SetFocus
  End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text7.SetFocus
  End If
End Sub

Private Sub Text60_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text61.SetFocus
  End If
End Sub

Private Sub Text61_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text62.SetFocus
  End If
End Sub

Private Sub Text62_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     ssPainel.Tab = 4
     Text63.SetFocus
  End If
End Sub

Private Sub Text63_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text64.SetFocus
  End If
End Sub

Private Sub Text64_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text65.SetFocus
  End If
End Sub

Private Sub Text65_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text93.SetFocus
  End If
End Sub

Private Sub Text66_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text67.SetFocus
  End If
End Sub

Private Sub Text67_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text68.SetFocus
  End If
End Sub

Private Sub Text68_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text69.SetFocus
  End If
End Sub

Private Sub Text69_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text70.SetFocus
  End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text8.SetFocus
  End If
End Sub

Private Sub Text70_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text71.SetFocus
  End If
End Sub

Private Sub Text71_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text72.SetFocus
  End If
End Sub

Private Sub Text72_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text73.SetFocus
  End If
End Sub

Private Sub Text73_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text76.SetFocus
  End If
End Sub

Private Sub Text74_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text75.SetFocus
  End If
End Sub

Private Sub Text75_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text77.SetFocus
  End If
End Sub

Private Sub Text76_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text74.SetFocus
  End If
End Sub

Private Sub Text77_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ssPainel.Tab = 5
     Text78.SetFocus
  End If
End Sub

Private Sub Text78_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text79.SetFocus
  End If
End Sub

Private Sub Text79_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text80.SetFocus
  End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text9.SetFocus
  End If
End Sub

Private Sub Text80_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text81.SetFocus
  End If
End Sub

Private Sub Text81_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text82.SetFocus
  End If
End Sub

Private Sub Text82_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text83.SetFocus
  End If
End Sub

Private Sub Text83_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text84.SetFocus
  End If
End Sub

Private Sub Text84_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text85.SetFocus
  End If
End Sub

Private Sub Text85_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text86.SetFocus
  End If
End Sub

Private Sub Text86_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text87 = DateAdd("m", Text85.Text, CDate(Text86.Text))
    Text86 = CDate(Text86)
    Text87.SetFocus
  End If
End Sub

Private Sub Text87_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text94.SetFocus
  End If
End Sub

Private Sub Text88_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Command6.SetFocus
  End If
End Sub

Private Sub Text89_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text6.SetFocus
  End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text10.SetFocus
  End If
End Sub

Private Sub Text90_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text21.SetFocus
  End If
End Sub

Private Sub Text91_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text36.SetFocus
  End If
End Sub

Private Sub Text92_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text51.SetFocus
  End If
End Sub

Private Sub Text93_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text66.SetFocus
  End If
End Sub

Private Function ConverteTexto()
Dim TxtBox As Object

For Each TxtBox In Me.Controls
    If TypeOf TxtBox Is TextBox Then
    TxtBox = StrConv(TxtBox, vbUpperCase)
End If
Next TxtBox
End Function

Private Sub Text94_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text88.SetFocus
  End If
End Sub

Private Sub Text95_Change()
If Text95 = "" Then
dtaContrato.RecordSource = "SELECT * FROM CONTRATO"
dtaContrato.Refresh
End If
End Sub

Private Sub Text95_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command2.SetFocus
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
