VERSION 5.00
Begin VB.Form frmClientess 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programa Imobiliária\Dados\Bdimobiliaria.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Clientes"
      Top             =   3120
      Width           =   3975
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   240
      TabIndex        =   41
      Top             =   4440
      Width           =   9015
      Begin VB.CommandButton Command5 
         Caption         =   "S&air"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7560
         TabIndex        =   46
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "P&esquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         TabIndex        =   45
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Excluir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   44
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Alterar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   43
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Novo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Picture         =   "Clientes3.frx":0000
         TabIndex        =   42
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados do Cliente:"
      Height          =   3015
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   9015
      Begin VB.TextBox Text14 
         DataField       =   "Site"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   5280
         TabIndex        =   30
         Text            =   "Text14"
         Top             =   2520
         Width           =   3615
      End
      Begin VB.TextBox Text13 
         DataField       =   "Email"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   720
         TabIndex        =   28
         Text            =   "Text13"
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox Text12 
         DataField       =   "TelCom"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   6480
         TabIndex        =   26
         Text            =   "Text12"
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox Text11 
         DataField       =   "TelRes"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Text            =   "Text11"
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox Text10 
         DataField       =   "Cpf"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   6480
         TabIndex        =   22
         Text            =   "Text10"
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox Text9 
         DataField       =   "RG"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   480
         TabIndex        =   20
         Text            =   "Text9"
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox Text8 
         DataField       =   "CepNoti"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   7080
         TabIndex        =   17
         Text            =   "Text8"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text7 
         DataField       =   "Estado"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   5400
         TabIndex        =   16
         Text            =   "Text7"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text6 
         DataField       =   "CidadeNoti"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   720
         TabIndex        =   14
         Text            =   "Text6"
         Top             =   1440
         Width           =   3975
      End
      Begin VB.TextBox Text5 
         DataField       =   "BairroNoti"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   6000
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         DataField       =   "RuaNoti"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Text            =   "Text4"
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox Text3 
         DataField       =   "Profissao"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Text            =   "Text3"
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         DataField       =   "Nacionalidade"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   6960
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         DataField       =   "Nome"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Site:"
         Height          =   195
         Left            =   4800
         TabIndex        =   29
         Top             =   2520
         Width           =   315
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "E-mail:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   2520
         Width           =   465
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Fone Comercial:"
         Height          =   195
         Left            =   5280
         TabIndex        =   25
         Top             =   2160
         Width           =   1140
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fone Residencial:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   1275
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "CPF:"
         Height          =   195
         Left            =   6000
         TabIndex        =   21
         Top             =   1800
         Width           =   345
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "RG:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   285
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cep:"
         Height          =   195
         Left            =   6600
         TabIndex        =   18
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Uf:"
         Height          =   195
         Left            =   5040
         TabIndex        =   15
         Top             =   1440
         Width           =   210
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         Height          =   195
         Left            =   5400
         TabIndex        =   12
         Top             =   1080
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Estado Civil:"
         Height          =   195
         Left            =   5520
         TabIndex        =   8
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Profissão:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nacionalidade:"
         Height          =   195
         Left            =   5760
         TabIndex        =   4
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Conjuge:"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   3240
      Width           =   9015
      Begin VB.TextBox Text19 
         DataField       =   "CPFConjuge"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   6960
         TabIndex        =   40
         Text            =   "Text19"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text18 
         DataField       =   "RGConjuge"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   4080
         TabIndex        =   38
         Text            =   "Text18"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text17 
         DataField       =   "ProfConjuge"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   960
         TabIndex        =   36
         Text            =   "Text17"
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox Text16 
         DataField       =   "NacionalConjuge"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   6960
         TabIndex        =   33
         Text            =   "Text16"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text15 
         DataField       =   "Conjuge"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   720
         TabIndex        =   32
         Text            =   "Text15"
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "CPF:"
         Height          =   195
         Left            =   6480
         TabIndex        =   39
         Top             =   720
         Width           =   345
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "RG:"
         Height          =   195
         Left            =   3600
         TabIndex        =   37
         Top             =   720
         Width           =   285
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Profissão:"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Nacionalidade:"
         Height          =   195
         Left            =   5760
         TabIndex        =   34
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programa Imobiliária\Dados\Bdimobiliaria.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Clientes"
      Top             =   5160
      Width           =   3975
   End
End
Attribute VB_Name = "frmClientess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
Text1.SetFocus
End Sub
