VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ALTERAR 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Alteração de Registro"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   4935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancelar"
      Height          =   540
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Alterar Lançamento"
      Height          =   540
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text1 
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
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1140
   End
   Begin VB.TextBox Text3 
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
      Height          =   315
      Left            =   3675
      TabIndex        =   1
      Top             =   750
      Width           =   1515
   End
   Begin VB.TextBox Text4 
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
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   4935
   End
   Begin MSMask.MaskEdBox Text2 
      Height          =   315
      Left            =   1950
      TabIndex        =   8
      Top             =   750
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16776960
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Locatário:"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Altere os dados necessários abaixo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   75
      TabIndex        =   7
      Top             =   75
      Width           =   3420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      Height          =   195
      Left            =   225
      TabIndex        =   6
      Top             =   525
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data de Vencimento:"
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   525
      Width           =   1500
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor do Recibo:"
      Height          =   195
      Left            =   3525
      TabIndex        =   4
      Top             =   525
      Width           =   1185
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Locador:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   630
   End
End
Attribute VB_Name = "ALTERAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

frmAluguel.Data1.Recordset.Edit
frmAluguel.Data1.Recordset.Fields("codigo") = Text1.Text
frmAluguel.Data1.Recordset.Fields("vencimento") = Text2.Text
frmAluguel.Data1.Recordset.Fields("valor") = Text3.Text
frmAluguel.Data1.Recordset.Fields("locador") = Text4.Text
frmAluguel.Data1.Recordset.Fields("locatario") = Text5.Text
frmAluguel.Data1.Recordset.Update
MsgBox ("Lançamento alterado com sucesso !!"), vbInformation, "Sucesso !!"
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Text1.Text = frmAluguel.Data1.Recordset.Fields("Codigo")
Text2.Text = frmAluguel.Data1.Recordset.Fields("Vencimento")
Text3.Text = frmAluguel.Data1.Recordset.Fields("Valor")
Text4.Text = frmAluguel.Data1.Recordset.Fields("Locador")
Text5.Text = frmAluguel.Data1.Recordset.Fields("Locatario")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub
