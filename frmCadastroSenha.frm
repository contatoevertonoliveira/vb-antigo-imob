VERSION 5.00
Begin VB.Form frmCadastroSenha 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3285
   ClientLeft      =   3720
   ClientTop       =   2775
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Gravar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cad&astrados"
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Novo"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Usuário:"
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   4215
      Begin VB.OptionButton optB 
         Caption         =   "Básico"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3240
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optI 
         Caption         =   "Intermediário"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optA 
         Caption         =   "Avançado"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Senha:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   2400
      TabIndex        =   14
      Top             =   1560
      Width           =   510
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   3840
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   3840
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nome / Usuário:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cadastro de Usuários  -  (Tela Administrador)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   3825
   End
End
Attribute VB_Name = "frmCadastroSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bd As DAO.Database
Dim Tb As DAO.Recordset

Private Sub cmdSair_Click()
If Command1.Enabled = False Then
    optA.Enabled = False
    optI.Enabled = False
    optB.Enabled = False
    Text1 = Empty
    Text1.Enabled = False
    Text2 = Empty
    Text2.Enabled = False
    Text3 = Empty
    Text3.Enabled = False
    Command1.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = False
Else
    If MsgBox("Quer sair do Cadastro de Usuários?", vbYesNo, "Sair do sistema") = vbYes Then
        Unload frmCadastroSenha
    Else
        Exit Sub
    End If
End If
End Sub

Private Sub Command1_Click()

If Tb.RecordCount = 0 Then
    Text1 = "301"
Else
    Tb.MoveLast
    Text1 = Tb!codigo + 1
End If
optA.Enabled = True
optA.Value = False
optI.Enabled = True
optI.Value = False
optB.Enabled = True
optB.Value = False
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Text2.Enabled = True
Text2 = Empty
Text2.SetFocus
Text3.Enabled = True
Text3 = Empty
Command1.Enabled = False
Command3.Enabled = False
Command4.Enabled = True

End Sub

Private Sub Command4_Click()

grava
MsgBox ("Usuário cadastrado com sucesso!!!")
Text1 = Empty
Text2 = Empty
Text3 = Empty
Command4.Enabled = False
Command1.Enabled = True
Command3.Enabled = True
optA.Enabled = False
optI.Enabled = False
optB.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False

End Sub

Private Sub Form_Load()
Set Bd = OpenDatabase("\\Maq5\c\Meus documentos\Documentos Backup\Programa Imobiliária\Dados\Bdimobiliaria.mdb")
Set Tb = Bd.OpenRecordset("Controle", dbOpenTable)
Tb.Index = "IndNome"
End Sub

Private Sub Text2_GotFocus()
Text2.BackColor = &HFFFF&
End Sub

Private Sub Text2_LostFocus()
Tb.Index = "IndNome"
Tb.Seek "=", Text2.Text
If Tb.NoMatch = False Then
    Text2 = Tb!nome
    MsgBox "Não pode existir usuários com o login iguais.", vbCritical
    MsgBox ("Preencha um novo ""Login""!")
    Text2.SetFocus
    Text2 = Empty
Else
    Text2.BackColor = &H80000005
End If
End Sub

Private Function grava()

Tb.AddNew
If Text1 <> "" Then Tb!codigo = Text1
If Text2 <> "" Then Tb!nome = Text2
If Text3 <> "" Then Tb!senha = Text3
If optA.Value = True Then
Tb!tipo = "Avançado"
End If
If optI.Value = True Then
Tb!tipo = "Intermediário"
End If
If optB.Value = True Then
Tb!tipo = "Básico"
End If

Tb.Update
End Function

Private Sub Text3_GotFocus()
Text3.BackColor = &HFFFF&
End Sub

Private Sub Text3_LostFocus()
Text3.BackColor = &H80000005
End Sub
