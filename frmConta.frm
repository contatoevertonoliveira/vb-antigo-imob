VERSION 5.00
Begin VB.Form frmConta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Conta"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "C&ancelar"
      Height          =   495
      Left            =   2760
      TabIndex        =   12
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cadastr&ar"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFC0&
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
         Left            =   960
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFFC0&
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
         Left            =   3000
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFC0&
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
         Left            =   960
         TabIndex        =   5
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFC0&
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
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFC0&
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
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Luz - R$:"
         Height          =   195
         Left            =   2280
         TabIndex        =   8
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Água - R$:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Locatário:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mês de Ref.:"
         Height          =   195
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmContas.Data3.Recordset("Codigo") = Text1
frmContas.Data3.Recordset("Conta") = Text2
frmContas.Data3.Recordset("Locatario") = Text3
frmContas.Data3.Recordset("Agua") = Text4
frmContas.Data3.Recordset("Luz") = Text5
frmContas.Data3.UpdateRecord
MsgBox ("Dados de contas cadastrados com sucesso!")
End Sub

Private Sub Command2_Click()
frmContas.Data3.Recordset.CancelUpdate
Unload Me
End Sub
