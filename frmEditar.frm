VERSION 5.00
Begin VB.Form frmEditar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editar Vencimento"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   3480
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   3480
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   3480
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "C&ancelar"
         Height          =   495
         Left            =   2880
         TabIndex        =   16
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Atualizar Dados"
         Default         =   -1  'True
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Line Line2 
         X1              =   1560
         X2              =   1800
         Y1              =   960
         Y2              =   1320
      End
      Begin VB.Line Line1 
         X1              =   1560
         X2              =   1800
         Y1              =   960
         Y2              =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Data Pagto.:"
         Height          =   195
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Multa:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Iptu:"
         Height          =   195
         Left            =   3360
         TabIndex        =   12
         Top             =   1560
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Valor Recebido:"
         Height          =   195
         Left            =   3360
         TabIndex        =   11
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor Proprietário:"
         Height          =   195
         Left            =   3360
         TabIndex        =   10
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pagto. Proprietário:"
         Height          =   195
         Left            =   1800
         TabIndex        =   9
         Top             =   840
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vencimento:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmEditar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If MsgBox("Tem certeza que deseja alterar esse vencimento?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then

If frmPrestacao.Visible = True Then
    frmPrestacao.Data1.Recordset.Edit
    frmPrestacao.Data1.Recordset("Vencimento") = frmEditar.Text1
    frmPrestacao.Data1.Recordset("Recebidos") = frmEditar.Text2
    frmPrestacao.Data1.Recordset("Valor") = frmEditar.Text3
    frmPrestacao.Data1.Recordset("Prop") = frmEditar.Text4
    frmPrestacao.Data1.Recordset("ValorProp") = frmEditar.Text5
    frmPrestacao.Data1.Recordset("Multa") = frmEditar.Text6
    frmPrestacao.Data1.UpdateRecord

    MsgBox ("Vencimento alterado com sucesso!")
    Unload Me
    
ElseIf frmAluguel.Visible = True Then
    frmAluguel.Data1.Recordset.Edit
    frmAluguel.Data1.Recordset("Vencimento") = frmEditar.Text1
    frmAluguel.Data1.Recordset("Recebidos") = frmEditar.Text2
    frmAluguel.Data1.Recordset("Valor") = frmEditar.Text3
    frmAluguel.Data1.Recordset("Prop") = frmEditar.Text4
    frmAluguel.Data1.Recordset("ValorProp") = frmEditar.Text5
    frmAluguel.Data1.Recordset("Multa") = frmEditar.Text6
    frmAluguel.Data1.Recordset("Iptu") = frmEditar.Text7
    frmAluguel.Data1.UpdateRecord

    MsgBox ("Vencimento alterado com sucesso!")
    Unload Me
End If
Else
    Exit Sub
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next

If frmPrestacao.Visible = True Then
If Command2.Caption = "C&ancelar" Then
    Text1 = frmPrestacao.Data1.Recordset("Vencimento")
    Text2 = frmPrestacao.Data1.Recordset("Recebidos")
    Text3 = frmPrestacao.Data1.Recordset("Valor")
    Text4 = frmPrestacao.Data1.Recordset("Prop")
    Text5 = frmPrestacao.Data1.Recordset("ValorProp")
    Text6 = frmPrestacao.Data1.Recordset("Multa")
    Text7 = frmPrestacao.Data1.Recordset("Iptu")
    Command2.Caption = "Sa&ir"
Else
    If MsgBox("Deseja Sair?", vbYesNo, "Sair da Edição") = vbYes Then
        Unload Me
    Else
        Exit Sub
    End If
End If
ElseIf frmAluguel.Visible = True Then
    If Command2.Caption = "C&ancelar" Then
    Text1 = frmAluguel.Data1.Recordset("Vencimento")
    Text2 = frmAluguel.Data1.Recordset("Recebidos")
    Text3 = frmAluguel.Data1.Recordset("Valor")
    Text4 = frmAluguel.Data1.Recordset("Prop")
    Text5 = frmAluguel.Data1.Recordset("ValorProp")
    Text6 = frmAluguel.Data1.Recordset("Multa")
    Text7 = frmAluguel.Data1.Recordset("Iptu")
    Command2.Caption = "Sa&ir"
Else
    If MsgBox("Deseja Sair?", vbYesNo, "Sair da Edição") = vbYes Then
        Unload Me
    Else
        Exit Sub
    End If
End If
End If
End Sub
