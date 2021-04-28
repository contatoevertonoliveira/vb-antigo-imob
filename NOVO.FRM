VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form NOVO 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Novo Lançamento"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   360
      TabIndex        =   12
      Top             =   2040
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Incluir Lançamento"
      Height          =   540
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancelar"
      Height          =   540
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   1815
   End
   Begin MSMask.MaskEdBox Text2 
      Height          =   315
      Left            =   1950
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   750
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
      TabIndex        =   2
      Top             =   750
      Width           =   1515
   End
   Begin VB.TextBox Text6 
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
      TabIndex        =   3
      Top             =   1440
      Width           =   4875
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   525
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data de Vencimento:"
      Height          =   195
      Left            =   1800
      TabIndex        =   8
      Top             =   525
      Width           =   1500
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor do Recibo:"
      Height          =   195
      Left            =   3525
      TabIndex        =   7
      Top             =   525
      Width           =   1185
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição:"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   765
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data de lançamento:"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preencha os dados abaixo para criar um novo registro:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   360
      TabIndex        =   4
      Top             =   75
      Width           =   4500
   End
End
Attribute VB_Name = "NOVO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim SITUACAO As String

If Option2.Value = True Then
    SITUACAO = ""
ElseIf Option1.Value = True Then
    SITUACAO = Date
End If

If Text5.Text = "Selecione ..." Then
    MsgBox ("Selecione o tipo de documento"), vbExclamation, "Tipo do documento ?"
Else
    
   frmAluguel.Data1.Recordset.AddNew
   frmAluguel.Data1.Recordset("nro_documento") = Text1.Text
   frmAluguel.Data1.Recordset("data_vencimento") = Text2.Text
   frmAluguel.Data1.Recordset("valor_documento") = Text3.Text
   frmAluguel.Data1.Recordset("cod_fornecedor") = Text4.Text
   frmAluguel.Data1.Recordset("tipo_documento") = Text5.Text
   frmAluguel.Data1.Recordset("Descrição") = Text6.Text
   frmAluguel.Data1.Recordset("data_lançamento") = Text7.Text
   frmAluguel.Data1.Recordset("data_pagamento") = SITUACAO
   frmAluguel.Data1.Recordset.Update
    MsgBox ("Lançamento incluído com sucesso !!"), vbInformation, "Sucesso !"
    Unload Me

End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text7.Text = Date
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub
