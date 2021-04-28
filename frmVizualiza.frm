VERSION 5.00
Begin VB.Form frmVizualiza 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Super Imob - Formulário Gera Contrato"
   ClientHeight    =   7470
   ClientLeft      =   2625
   ClientTop       =   165
   ClientWidth     =   9990
   Icon            =   "frmVizualiza.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Caption         =   "Mais Informações:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   6000
      TabIndex        =   35
      Top             =   4560
      Width           =   3615
      Begin VB.TextBox text 
         Height          =   615
         Index           =   7
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox text 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   8
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   37
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox text 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   9
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   36
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observação:"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Assinatura:"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   1200
         Width           =   1170
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   "Início Contr.:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   1560
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Informações Adicionais:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   6000
      TabIndex        =   18
      Top             =   1680
      Width           =   3615
      Begin VB.TextBox text 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox text 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   2415
      End
      Begin VB.ComboBox cboUf 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox text 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox text 
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   22
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox text 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox text 
         Height          =   285
         Index           =   6
         Left            =   2400
         TabIndex        =   20
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox text 
         Height          =   285
         Index           =   5
         Left            =   1080
         TabIndex        =   19
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "End. Imóvel Alugado:"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         Height          =   195
         Left            =   2640
         TabIndex        =   33
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "Valor R$:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   660
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "Finalidade:"
         Height          =   195
         Left            =   1800
         TabIndex        =   30
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         Caption         =   "Índice"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "Reajuste:"
         Height          =   195
         Left            =   2400
         TabIndex        =   28
         Top             =   2040
         Width           =   675
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         Caption         =   "Prazo Locação:"
         Height          =   195
         Left            =   1080
         TabIndex        =   27
         Top             =   2040
         Width           =   1125
      End
   End
   Begin VB.ComboBox cboLoc 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1200
      Width           =   3135
   End
   Begin VB.ComboBox cboProp 
      Height          =   315
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Limpar"
      Enabled         =   0   'False
      Height          =   735
      Left            =   4920
      Picture         =   "frmVizualiza.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Adicion&ar"
      Enabled         =   0   'False
      Height          =   735
      Left            =   6000
      Picture         =   "frmVizualiza.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sa&ir"
      Height          =   735
      Left            =   8400
      Picture         =   "frmVizualiza.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ger&ar"
      Enabled         =   0   'False
      Height          =   735
      Left            =   7200
      Picture         =   "frmVizualiza.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cadastrar e Gerar Contrato"
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   84
      Top             =   240
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Locador..............:"
      Height          =   195
      Left            =   240
      TabIndex        =   83
      Top             =   720
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nacionalidade....:"
      Height          =   195
      Left            =   240
      TabIndex        =   82
      Top             =   960
      Width           =   1245
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado Civil........:"
      Height          =   195
      Left            =   240
      TabIndex        =   81
      Top             =   1200
      Width           =   1230
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Profissão............:"
      Height          =   195
      Left            =   240
      TabIndex        =   80
      Top             =   1440
      Width           =   1230
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cpf.....................:"
      Height          =   195
      Left            =   240
      TabIndex        =   79
      Top             =   1680
      Width           =   1230
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rg......................:"
      Height          =   195
      Left            =   240
      TabIndex        =   78
      Top             =   1920
      Width           =   1245
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço...........:"
      Height          =   195
      Left            =   240
      TabIndex        =   77
      Top             =   2160
      Width           =   1230
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro.................:"
      Height          =   195
      Left            =   240
      TabIndex        =   76
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Locatário...........:"
      Height          =   195
      Left            =   240
      TabIndex        =   75
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nacionalidade...:"
      Height          =   195
      Left            =   240
      TabIndex        =   74
      Top             =   3000
      Width           =   1200
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado Civil.......:"
      Height          =   195
      Left            =   240
      TabIndex        =   73
      Top             =   3240
      Width           =   1185
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Profissão...........:"
      Height          =   195
      Left            =   240
      TabIndex        =   72
      Top             =   3480
      Width           =   1185
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cpf....................:"
      Height          =   195
      Left            =   240
      TabIndex        =   71
      Top             =   3720
      Width           =   1185
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rg.....................:"
      Height          =   195
      Left            =   240
      TabIndex        =   70
      Top             =   3960
      Width           =   1200
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço..........:"
      Height          =   195
      Left            =   240
      TabIndex        =   69
      Top             =   4200
      Width           =   1185
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro................:"
      Height          =   195
      Left            =   240
      TabIndex        =   68
      Top             =   4440
      Width           =   1170
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço..........:"
      Height          =   195
      Left            =   240
      TabIndex        =   67
      Top             =   4800
      Width           =   1185
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro................:"
      Height          =   195
      Left            =   240
      TabIndex        =   66
      Top             =   5040
      Width           =   1170
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Finalidade.........:"
      Height          =   195
      Left            =   240
      TabIndex        =   65
      Top             =   5280
      Width           =   1170
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aluguél.............:"
      Height          =   195
      Left            =   240
      TabIndex        =   64
      Top             =   5640
      Width           =   1155
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label31"
      DataField       =   "Nome"
      DataSource      =   "Data1"
      Height          =   195
      Left            =   1560
      TabIndex        =   63
      Top             =   720
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label32"
      DataField       =   "Nacionalidade"
      DataSource      =   "Data1"
      Height          =   195
      Left            =   1560
      TabIndex        =   62
      Top             =   960
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label33"
      DataField       =   "EstadoCivil"
      DataSource      =   "Data1"
      Height          =   195
      Left            =   1560
      TabIndex        =   61
      Top             =   1200
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label34"
      DataField       =   "Profissao"
      DataSource      =   "Data1"
      Height          =   195
      Left            =   1560
      TabIndex        =   60
      Top             =   1440
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label35"
      DataField       =   "CPF"
      DataSource      =   "Data1"
      Height          =   195
      Left            =   1560
      TabIndex        =   59
      Top             =   1680
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label36"
      DataField       =   "RG"
      DataSource      =   "Data1"
      Height          =   195
      Left            =   1560
      TabIndex        =   58
      Top             =   1920
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label37"
      DataField       =   "RuaNoti"
      DataSource      =   "Data1"
      Height          =   195
      Left            =   1560
      TabIndex        =   57
      Top             =   2160
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label38"
      DataField       =   "BairroNoti"
      DataSource      =   "Data1"
      Height          =   195
      Left            =   1560
      TabIndex        =   56
      Top             =   2400
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label39"
      Height          =   195
      Left            =   1560
      TabIndex        =   55
      Top             =   2760
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label40"
      Height          =   195
      Left            =   1560
      TabIndex        =   54
      Top             =   3000
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label41"
      Height          =   195
      Left            =   1560
      TabIndex        =   53
      Top             =   3240
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label42"
      Height          =   195
      Left            =   1560
      TabIndex        =   52
      Top             =   3480
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label43"
      Height          =   195
      Left            =   1560
      TabIndex        =   51
      Top             =   3720
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label44"
      Height          =   195
      Left            =   1560
      TabIndex        =   50
      Top             =   3960
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label45 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label45"
      Height          =   195
      Left            =   1560
      TabIndex        =   49
      Top             =   4200
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label46 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label46"
      Height          =   195
      Left            =   1560
      TabIndex        =   48
      Top             =   4440
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label47 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label47"
      Height          =   195
      Left            =   1560
      TabIndex        =   47
      Top             =   4800
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label48"
      Height          =   195
      Left            =   1560
      TabIndex        =   46
      Top             =   5040
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label49 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label49"
      Height          =   195
      Left            =   1560
      TabIndex        =   45
      Top             =   5280
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label50"
      Height          =   195
      Left            =   1560
      TabIndex        =   44
      Top             =   5640
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proprietário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   6480
      TabIndex        =   43
      Top             =   360
      Width           =   990
   End
   Begin VB.Label Label65 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Locatário:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   6480
      TabIndex        =   42
      Top             =   960
      Width           =   870
   End
   Begin VB.Label Label56 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label56"
      Height          =   195
      Left            =   1560
      TabIndex        =   15
      Top             =   7080
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label55 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label55"
      Height          =   195
      Left            =   1560
      TabIndex        =   14
      Top             =   6840
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label54 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label54"
      Height          =   195
      Left            =   1560
      TabIndex        =   13
      Top             =   6600
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label53 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label53"
      Height          =   195
      Left            =   1560
      TabIndex        =   12
      Top             =   6360
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label52 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label52"
      Height          =   195
      Left            =   1560
      TabIndex        =   11
      Top             =   6120
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label51 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label51"
      Height          =   195
      Left            =   1560
      TabIndex        =   10
      Top             =   5880
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Final..................:"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   7080
      Width           =   1185
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ínicio.................:"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   6840
      Width           =   1185
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prazo Locação.:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   6600
      Width           =   1170
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reajuste...........:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   6360
      Width           =   1170
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Índice...............:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   6120
      Width           =   1155
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D. de Pagto......:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   5880
      Width           =   1170
   End
End
Attribute VB_Name = "frmVizualiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wrdApp As Word.Application
Dim wrdAppF As Word.Application
Dim wrdSelection As Word.Selection
Dim Nome_Arq As String
Dim Meuerro$
Dim X As Integer
Dim Bd As DAO.Database
Dim Rs3 As DAO.Recordset

Private Sub cboLoc_Click()
HabilitaLabels2
Set Rs2 = Nothing
Rs2.Open "Select Nome, Nacionalidade, Profissao, EstadoCivil, Cpf, Rg, Ruanoti, Bairronoti From Loc where Nome='" & cboLoc.Text & "'", CN, adOpenStatic, adLockPessimistic
CarregaLabelLoc
Command3.Enabled = True
Command1.Enabled = True
cmdSair.Caption = "C&ancelar"
AbreCaixas
Text(0).SetFocus
End Sub

Private Sub cboProp_Click()
HabilitaLabels
Set Rs = Nothing
Rs.Open "Select Nome, Nacionalidade, Profissao, EstadoCivil, Cpf, Rg, Ruanoti, Bairronoti From Prop where Nome='" & cboProp.Text & "'", CN, adOpenStatic, adLockPessimistic
CarregaLabelProp
Command3.Enabled = True
Command1.Enabled = True
cmdSair.Caption = "C&ancelar"
cboLoc.Enabled = True
End Sub

Private Sub cmdInfo_Click()

Label47.Caption = Text(0).Text
Label48.Caption = Text(1).Text & " - " & cboUf.Text
Label49.Caption = Text(3).Text
Label50.Caption = Text(2).Text
Label51.Caption = "Todo dia " & Text(8)
Label52.Caption = Text(4).Text
Label53.Caption = Text(6).Text
Label54.Caption = Text(5).Text & " " & "meses"
Label55.Caption = Text(9).Text
Label56.Caption = DateAdd("m", Text(5).Text, CDate(Text(9).Text))
cmdInfo.Enabled = False
Command3.Enabled = False
LimpaTudo

End Sub

Private Sub cmdSair_Click()

If cmdSair.Caption = "C&ancelar" Then
LimpaTudo
LimpaLabels
FechaCaixas
cboProp.ListIndex = -1
cboLoc.ListIndex = -1
cmdSair.Caption = "Sa&ir"
Command3.Enabled = False
Command1.Enabled = False
cmdInfo.Enabled = False
If cboProp.ListIndex = -1 Then
cboLoc.Enabled = False
End If
If cboLoc.ListIndex = -1 Then
FechaCaixas
Else
AbreCaixas
Text(0).SetFocus
End If
Else
If MsgBox("Quer sair do Gera Contratos?", vbYesNo, "Sair do Gerador") = vbYes Then
    Unload Me
    CN.Close
Else
    Exit Sub
End If
End If
End Sub

Private Sub Command1_Click()

If MsgBox("Você deseja Gerar o Contrato?", vbYesNo, " || Cadastro ||Gerador ||") = vbYes Then
PreencheContrato
GravaDados
LimpaTudo
FechaCaixas
Command3.Enabled = False
Command1.Enabled = False
cmdSair.Caption = "Sa&ir"
cmdInfo.Enabled = False
MsgBox ("O Contrato foi preenchido e os dados cadastrados com sucesso!")
MsgBox ("Contrato cadastrado com sucesso! - O Novo Código cadastrado é: " & "KC" & Format(Rs3.RecordCount, "000"))
Bd.Close
LimpaTudo
LimpaLabels
FechaCaixas
cboProp.ListIndex = -1
cboLoc.ListIndex = -1
Command3.Enabled = False
Command1.Enabled = False
cmdSair.Caption = "Sa&ir"
cmdInfo.Enabled = False
If cboProp.ListIndex = -1 Then
cboLoc.Enabled = False
End If
If cboLoc.ListIndex = -1 Then
FechaCaixas
End If
Else
GravaDados
LimpaTudo
FechaCaixas
Command3.Enabled = False
Command1.Enabled = False
cmdSair.Caption = "Sa&ir"
cmdInfo.Enabled = False
MsgBox ("Contrato cadastrado com sucesso! - O Novo Código cadastrado é: " & "KC" & Format(Rs3.RecordCount, "000"))
cmdSair.Caption = "Sa&ir"
Bd.Close
LimpaTudo
LimpaLabels
FechaCaixas
cboProp.ListIndex = -1
cboLoc.ListIndex = -1
Command3.Enabled = False
Command1.Enabled = False
cmdSair.Caption = "Sa&ir"
cmdInfo.Enabled = False
If cboProp.ListIndex = -1 Then
cboLoc.Enabled = False
End If
If cboLoc.ListIndex = -1 Then
FechaCaixas
End If
End If
End Sub

Private Sub Command3_Click()
LimpaTudo
LimpaLabels
FechaCaixas
cboProp.ListIndex = -1
cboLoc.ListIndex = -1
Command3.Enabled = False
Command1.Enabled = False
cmdSair.Caption = "Sa&ir"
cmdInfo.Enabled = False
If cboProp.ListIndex = -1 Then
cboLoc.Enabled = False
End If
If cboLoc.ListIndex = -1 Then
FechaCaixas
Else
AbreCaixas
Text(0).SetFocus
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Erro

If KeyCode = 40 Then
End If
If KeyCode = vbKeyLeft And ActiveControl.SelStart = 0 Then
        SendKeys "+{tab}"
    ElseIf KeyCode = vbKeyRight And ActiveControl.SelStart = Len(ActiveControl.Text) Then
        SendKeys "{tab}"
    End If
Erro:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
     SendKeys vbTab
     KeyAscii = 0
  End If
 If KeyAscii = 27 Then
 If MsgBox("Quer sair do Cadastro e Gerador de Contrato?", vbYesNo, "Sair do Gerador") = vbYes Then
    Unload frmVizualiza
    CN.Close
Else
    Exit Sub
End If
End If
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Height) / 2
Me.Left = (Screen.Width - Width) / 2
Conecta
CarregaCombo1
CarregaCombo2
CarregaUf
Label1.Caption = "CONTRATO DE  LOCAÇÃO"
FechaCaixas
End Sub

Private Function CarregaCombo1()
cboProp.Clear
    Do While Not Rs.EOF
        With cboProp
            .AddItem Rs!nome
        End With
Rs.MoveNext
    Loop
Rs.Close
End Function

Private Function CarregaCombo2()
cboLoc.Clear
    Do While Not Rs2.EOF
        With cboLoc
            .AddItem Rs2!nome
        End With
Rs2.MoveNext
    Loop
Rs2.Close
End Function

Private Function HabilitaLabels()

Label31.Visible = True
Label32.Visible = True
Label33.Visible = True
Label34.Visible = True
Label35.Visible = True
Label36.Visible = True
Label37.Visible = True
Label38.Visible = True

End Function

Private Function CarregaLabelProp()

Label31.Caption = Rs!nome
Label32.Caption = Rs!nacionalidade
Label33.Caption = Rs!Estadocivil
Label34.Caption = Rs!profissao
Label35.Caption = Rs!cpf
Label36.Caption = Rs!rg
Label37.Caption = Rs!ruanoti
Label38.Caption = Rs!bairronoti

End Function

Private Function CarregaLabelLoc()

If Rs2!nome <> "" Then Label39.Caption = Rs2!nome
If Rs2!nacionalidade <> "" Then Label40.Caption = Rs2!nacionalidade
If Rs2!Estadocivil <> "" Then Label41.Caption = Rs2!Estadocivil
If Rs2!profissao <> "" Then Label42.Caption = Rs2!profissao
If Rs2!cpf <> "" Then Label43.Caption = Rs2!cpf
If Rs2!rg <> "" Then Label44.Caption = Rs2!rg
If Rs2!ruanoti <> "" Then Label45.Caption = Rs2!ruanoti
If Rs2!bairronoti <> "" Then Label46.Caption = Rs2!bairronoti

End Function

Private Function CarregaUf()

With cboUf
        .AddItem "AC"
        .AddItem "AL"
        .AddItem "AM"
        .AddItem "AP"
        .AddItem "BA"
        .AddItem "CE"
        .AddItem "DF"
        .AddItem "ES"
        .AddItem "GO"
        .AddItem "MA"
        .AddItem "MG"
        .AddItem "MS"
        .AddItem "MT"
        .AddItem "PA"
        .AddItem "PB"
        .AddItem "PE"
        .AddItem "PI"
        .AddItem "PR"
        .AddItem "RJ"
        .AddItem "RN"
        .AddItem "RO"
        .AddItem "RR"
        .AddItem "Rs"
        .AddItem "SC"
        .AddItem "SE"
        .AddItem "SP"
        .AddItem "TO"
End With
End Function

Private Function HabilitaLabels2()

Label39.Visible = True
Label40.Visible = True
Label41.Visible = True
Label42.Visible = True
Label43.Visible = True
Label44.Visible = True
Label45.Visible = True
Label46.Visible = True

Label47.Visible = True
Label47.Caption = ""
Label48.Visible = True
Label48.Caption = ""
Label49.Visible = True
Label49.Caption = ""
Label50.Visible = True
Label50.Caption = ""
Label51.Visible = True
Label51.Caption = ""
Label52.Visible = True
Label52.Caption = ""
Label53.Visible = True
Label53.Caption = ""
Label54.Visible = True
Label54.Caption = ""
Label55.Visible = True
Label55.Caption = ""
Label56.Visible = True
Label56.Caption = ""

End Function

Private Function LimpaTudo()
Dim Txt_Objeto As Object

For Each Txt_Objeto In Me.Controls
If TypeOf Txt_Objeto Is TextBox Then
    Txt_Objeto.Text = Empty
End If
Next Txt_Objeto
cboUf.ListIndex = -1
End Function

Private Function FechaCaixas()

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
cboUf.Enabled = False

End Function

Private Function AbreCaixas()

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
cboUf.Enabled = True

End Function

Private Sub Substitui_Var1(Header As String, data As String, oWord As Object)
    On Error Resume Next
    With oWord.Selection.Find
        .ClearFormatting
        .Text = Header
        .Execute Forward:=True
    End With
    Clipboard.Clear
    Clipboard.SetText (data)
    oWord.Selection.Paste
    Clipboard.Clear
End Sub

Private Function PreencheContrato()
Dim ObjWord As New Word.Application
    
Me.MousePointer = 11
ObjWord.Visible = False
    
    If MsgBox("Deseja Autopreencher um Contrato de Locação com Dados do Proprietário e Cliente selecionados?", vbQuestion + vbYesNo) = vbYes Then
    
    CopyFile "C:\Programa Imobiliária\Contratos\Contrato.DOC", "C:\Programa Imobiliária\Contratos\" & Label31.Caption & ".doc"
    ObjWord.Documents.Open ("C:\Programa Imobiliária\Contratos\" & Label31.Caption & ".doc")
    
    Call Substitui_Var1("@Locador", Label31.Caption, ObjWord)
    Call Substitui_Var1("@Nacionalidade", Label32.Caption, ObjWord)
    Call Substitui_Var1("@EstadocivilLocador", Label33.Caption, ObjWord)
    Call Substitui_Var1("@ProfissaoLocador", Label34.Caption, ObjWord)
    Call Substitui_Var1("@CpfLocador", Label35.Caption, ObjWord)
    Call Substitui_Var1("@RgLocador", Label36.Caption, ObjWord)
    Call Substitui_Var1("@EndLocador", Label37.Caption, ObjWord)
    Call Substitui_Var1("@BairroLocador", Label38.Caption, ObjWord)
    
    Call Substitui_Var1("@Locatario", Label39.Caption, ObjWord)
    Call Substitui_Var1("@NacionalidadeLocatario", Label40.Caption, ObjWord)
    Call Substitui_Var1("@EstadocivilLocatario", Label41.Caption, ObjWord)
    Call Substitui_Var1("@ProfissaoLocatario", Label42.Caption, ObjWord)
    Call Substitui_Var1("@CpfLocatario", Label43.Caption, ObjWord)
    Call Substitui_Var1("@RgLocatario", Label44.Caption, ObjWord)
    Call Substitui_Var1("@EndLocatario", Label45.Caption, ObjWord)
    Call Substitui_Var1("@BairroLocatario", Label46.Caption, ObjWord)
    
    Call Substitui_Var1("@ImovelAlugar", Label47.Caption, ObjWord)
    Call Substitui_Var1("@ImovelBairro", Label48.Caption, ObjWord)
    Call Substitui_Var1("@Finalidade", Label49.Caption, ObjWord)
    
    Call Substitui_Var1("@Aluguel", Label50.Caption, ObjWord)
    Call Substitui_Var1("@Dia", Label51.Caption, ObjWord)
    Call Substitui_Var1("@IGPM/FGV", Label52.Caption, ObjWord)
    Call Substitui_Var1("@Anual", Label53.Caption, ObjWord)
    Call Substitui_Var1("@DataMeses", Label54.Caption, ObjWord)
    Call Substitui_Var1("@Inicio", Label55.Caption, ObjWord)
    Call Substitui_Var1("@Final", Label56.Caption, ObjWord)
    
    Call Substitui_Var1("@Obs", Text(7), ObjWord)
    Call Substitui_Var1("@Data2", Label55.Caption, ObjWord)
    Call Substitui_Var1("@AssinaturaLocador", Label31.Caption, ObjWord)
    Call Substitui_Var1("@AssinaturaLocatario", Label39.Caption, ObjWord)
    Call Substitui_Var1("@Testemunha1", "Testemunha", ObjWord)
    Call Substitui_Var1("@Testemunha2", "Testemunha", ObjWord)
    
    ObjWord.ActiveDocument.Save
    ObjWord.Quit
    Set ObjWord = Nothing
    Me.MousePointer = 0
    MsgBox "Contrato Gerado com Sucesso em: " & vbCrLf & App.Path & "\Contrato\" & Label31.Caption & ".doc", vbInformation, " Contrato Gerado "
    End If
End Function

Private Function GravaDados()

Set Bd = OpenDatabase(App.Path & "\dados\bdimobiliaria")
Set Rs3 = Bd.OpenRecordset("Contrato", dbOpenTable)

Rs3.AddNew

Rs3!codigo = "KC" & Format(Rs3.RecordCount + 1, "000")
If Label31.Caption <> "" Then Rs3("Locador") = Label31.Caption
If Label32.Caption <> "" Then Rs3("NLocador") = Label32.Caption
If Label33.Caption <> "" Then Rs3("ELocador") = Label33.Caption
If Label34.Caption <> "" Then Rs3("ProfLocador") = Label34.Caption
If Label35.Caption <> "" Then Rs3("CpfLocador") = Label35.Caption
If Label36.Caption <> "" Then Rs3("RgLocador") = Label36.Caption
If Label37.Caption <> "" Then Rs3("EndLocador") = Label37.Caption
If Label38.Caption <> "" Then Rs3("BLocador") = Label38.Caption

If Label39.Caption <> "" Then Rs3("Locatario") = Label39.Caption
If Label40.Caption <> "" Then Rs3("NLocatario") = Label40.Caption
If Label41.Caption <> "" Then Rs3("ELocatario") = Label41.Caption
If Label42.Caption <> "" Then Rs3("ProfLocatario") = Label42.Caption
If Label43.Caption <> "" Then Rs3("CpfLocatario") = Label43.Caption
If Label44.Caption <> "" Then Rs3("RgLocatario") = Label44.Caption
If Label45.Caption <> "" Then Rs3("EndLocatario") = Label45.Caption
If Label46.Caption <> "" Then Rs3("BLocatario") = Label46.Caption

If Label47.Caption <> "" Then Rs3("ImovelLocado") = Label47.Caption
If Label48.Caption <> "" Then Rs3("BImovel") = Label48.Caption
If Label49.Caption <> "" Then Rs3("Finalidade") = Label49.Caption

If Label50.Caption <> "" Then Rs3("Aluguel") = Label50.Caption
If Label51.Caption <> "" Then Rs3("Dia") = Label51.Caption
If Label52.Caption <> "" Then Rs3("Indice") = Label52.Caption
If Label53.Caption <> "" Then Rs3("Periodo") = Label53.Caption
If Label54.Caption <> "" Then Rs3("Prazo") = Label54.Caption
If Label55.Caption <> "" Then Rs3("Inicio") = Label55.Caption
If Label56.Caption <> "" Then Rs3("Final") = Label56.Caption
If Label55.Caption <> "" Then Rs3("Dia2") = Label55.Caption
If Text(7) <> "" Then Rs3("Obs") = Text(7)

Rs3.Update

End Function

Private Function LimpaLabels()

Label31 = ""
Label32 = ""
Label33 = ""
Label34 = ""
Label35 = ""
Label36 = ""
Label37 = ""
Label38 = ""
Label39 = ""
Label40 = ""
Label41 = ""
Label42 = ""
Label43 = ""
Label44 = ""
Label45 = ""
Label46 = ""
Label47.Caption = ""
Label48.Caption = ""
Label49.Caption = ""
Label50.Caption = ""
Label51.Caption = ""
Label52.Caption = ""
Label53.Caption = ""
Label54.Caption = ""
Label55.Caption = ""
Label56.Caption = ""

End Function

Private Sub Text_Change(Index As Integer)
On Error Resume Next

Select Case Index
Case 8
    If Len(Text(8)) = 2 Then
    Text(8) = Text(8) + "/"
    Text(8).SelStart = 3
    End If

    If Len(Text(8)) = 5 Then
    Text(8) = Text(8) + "/"
    Text(8).SelStart = 6
    End If

Case 9
    If Len(Text(9)) = 2 Then
    Text(9) = Text(9) + "/"
    Text(9).SelStart = 3
    End If

    If Len(Text(9)) = 5 Then
    Text(9) = Text(9) + "/"
    Text(9).SelStart = 6
    End If
If Len(Text(9).Text) = 8 Then
    cmdInfo.Enabled = True
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

End Select
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 2
Select Case KeyAscii
        Case Is = 8
        Case 48 To 57
        Case Else
        KeyAscii = 0
    End Select
Case 5
Select Case KeyAscii
        Case Is = 8
        Case 48 To 57
        Case Else
        KeyAscii = 0
    End Select
End Select
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
Text(2) = Format(Text(2), "currency")

Case 3
Text(3).BackColor = &H80000005
Text(3) = StrConv(Text(3), vbUpperCase)

Case 4
Text(4).BackColor = &H80000005
Text(4) = StrConv(Text(4), vbUpperCase)

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
Text(8) = CDate(Text(8).Text)

Case 9
    Text(9).BackColor = &H80000005
If Len(Text(9).Text) = 8 Then
    Text(9) = CDate(Text(9).Text)
Else
    MsgBox ("Preencha uma data correta!")
    Text(9).SetFocus
End If
End Select
End Sub
