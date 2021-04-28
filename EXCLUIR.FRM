VERSION 5.00
Begin VB.Form EXCLUIR 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Excluir lançamento"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancelar"
      Height          =   540
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Excluir Lançamento"
      Height          =   540
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox Text5 
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
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2040
      Width           =   4875
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
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1440
      Width           =   4845
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
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   750
      Width           =   1515
   End
   Begin VB.TextBox Text2 
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
      Left            =   1950
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   750
      Width           =   1140
   End
   Begin VB.TextBox Text1 
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
      Left            =   375
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   750
      Width           =   1140
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Locatário:"
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   1800
      Width           =   705
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Locador:"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor do Recibo:"
      Height          =   195
      Left            =   3525
      TabIndex        =   5
      Top             =   525
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data de Vencimento:"
      Height          =   195
      Left            =   1800
      TabIndex        =   3
      Top             =   525
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   525
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Você tem certeza que deseja excluir o lançamento com os dados abaixo?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5340
   End
End
Attribute VB_Name = "EXCLUIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If frmAluguel.Visible = True Then
frmAluguel.Data1.Recordset.Delete
MsgBox ("O Registro foi excluido com sucesso !!"), vbInformation, "Sucesso !!"
Unload Me
ElseIf frmPrestacao.Visible = True Then
frmPrestacao.Data1.Recordset.Delete
MsgBox ("O Registro foi excluido com sucesso !!"), vbInformation, "Sucesso !!"
Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
If frmAluguel.Visible = True Then
Text1.Text = frmAluguel.Data1.Recordset.Fields("codigo")
Text2.Text = frmAluguel.Data1.Recordset.Fields("vencimento")
Text3.Text = frmAluguel.Data1.Recordset.Fields("valor")
Text4.Text = frmAluguel.Data1.Recordset.Fields("locador")
Text5.Text = frmAluguel.Data1.Recordset.Fields("locatario")
ElseIf frmPrestacao.Visible = True Then
Label6 = "Vendedor"
Label7 = "Comprador"
Text1.Text = frmPrestacao.Data1.Recordset.Fields("codigo")
Text2.Text = frmPrestacao.Data1.Recordset.Fields("vencimento")
Text3.Text = frmPrestacao.Data1.Recordset.Fields("valor")
Text4.Text = frmPrestacao.Data1.Recordset.Fields("VENDEDOR")
Text5.Text = frmPrestacao.Data1.Recordset.Fields("COMPRADOR")
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub
