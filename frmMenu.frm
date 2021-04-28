VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CADASTROS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CONTRATOS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "OPÇÕES:"
      ForeColor       =   &H8000000E&
      Height          =   4335
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   8175
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OPÇÕES"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1800
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
