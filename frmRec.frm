VERSION 5.00
Begin VB.Form frmRec 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Geração de Recibos"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Dados para a geração de novos recibos:"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton Command2 
         Caption         =   "C&ancelar"
         Height          =   615
         Left            =   2400
         TabIndex        =   10
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ger&ar + Recibos"
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
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
         Left            =   3120
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
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
         Left            =   1080
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
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
         Left            =   1080
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
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
         Left            =   1080
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor R$:"
         Height          =   195
         Left            =   3240
         TabIndex        =   7
         Top             =   720
         Width           =   660
      End
      Begin VB.Line Line2 
         X1              =   3000
         X2              =   2520
         Y1              =   960
         Y2              =   1440
      End
      Begin VB.Line Line1 
         X1              =   2040
         X2              =   3000
         Y1              =   480
         Y2              =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Término:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Início:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Mes As Integer
Dim ANO As Integer
Dim Data As String
Dim Valor As String

Mes = Format(Text2, "mm")
ANO = Format(Text2, "yy")
Valor = Val(frmRecibos.Label30)

For i = 1 To Val(Text1.Text)
    Mes = Mes + 1
    Valor = Valor + 1
    If Mes > 12 Then
       Mes = 1
       ANO = ANO + 1
   End If
    
  dia1 = Format(Text2, "dd")
  Dia = Verifica_dia(dia1, Mes)
  Data = Dia & "/" & Format(Mes, "00") & "/" & Format(ANO, "00")
  
  frmVencimentos.Data2.Recordset.AddNew
  frmVencimentos.Data2.Recordset.Fields(1) = CLng(frmVencimentos.Text7)
  frmVencimentos.Data2.Recordset.Fields(2) = Format(CDate(Data), "dd/mm/yy")
  frmVencimentos.Data2.Recordset.Fields(3) = frmVencimentos.Text1
  frmVencimentos.Data2.Recordset.Fields(4) = frmVencimentos.Text2
  frmVencimentos.Data2.Recordset.Fields(5) = frmVencimentos.Text4
  frmVencimentos.Data2.Recordset.Fields(6) = Text4
  frmVencimentos.Data2.Recordset.Fields(7) = frmVencimentos.Label12
  frmVencimentos.Data2.Recordset.Fields(8) = frmVencimentos.Label14
  frmVencimentos.Data2.Recordset.Fields(11) = Val(Valor)
  frmVencimentos.Data2.Recordset.Update
Next
MsgBox ("Recibos Gerados e adicionados com sucesso!")
frmVencimentos.Data2.RecordSource = "SELECT * FROM VENCIMENTOS WHERE LOCATARIO Like '" & frmVencimentos.Text2 & "*'"
frmVencimentos.Data2.Refresh
Unload Me
End Sub


Public Function Verifica_dia(Dia, Mes)
Dim diasDoMes As Variant

Dia = Val(Dia)

diasDoMes = Array(31, 28, 30, 30, 31, 30, 31, 30, 30, 31, 30, 31)

If Dia = 31 Then
Verifica_dia = diasDoMes(Mes - 1)
Else
Verifica_dia = Dia
End If

End Function

Private Sub Command2_Click()
Unload Me
End Sub
