VERSION 5.00
Begin VB.Form frmClientes 
   Caption         =   "Agenda de Telefones e Endereços"
   ClientHeight    =   5325
   ClientLeft      =   1245
   ClientTop       =   1605
   ClientWidth     =   9480
   Icon            =   "frmClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   9480
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   240
      TabIndex        =   23
      Top             =   4320
      Width           =   9135
      Begin VB.CommandButton cmda 
         Caption         =   "<<<<"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdp 
         Caption         =   ">>>>"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8040
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6720
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5280
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3840
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdAlterar 
         Caption         =   "&Alterar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar Formulário"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7200
      TabIndex        =   22
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   615
      Left            =   480
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3360
      Width           =   8655
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   8
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      MaxLength       =   50
      TabIndex        =   6
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4560
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2160
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7200
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1560
      Width           =   6615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   285
      Left            =   480
      TabIndex        =   19
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   240
      TabIndex        =   18
      Top             =   480
      Width           =   9135
      Begin VB.CommandButton cmdBusc 
         Caption         =   "&Buscar Cliente"
         Height          =   375
         Left            =   3120
         TabIndex        =   35
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Text11 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6840
         MaxLength       =   10
         TabIndex        =   9
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Text10 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7560
         MaxLength       =   9
         TabIndex        =   5
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   7
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5040
         TabIndex        =   24
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Observação:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   2640
         Width           =   915
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fone Comercial:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   6840
         TabIndex        =   33
         Top             =   2040
         Width           =   1140
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Fone Residencial:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   4440
         TabIndex        =   32
         Top             =   2040
         Width           =   1275
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cep:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   7560
         TabIndex        =   31
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3360
         TabIndex        =   30
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   4320
         TabIndex        =   28
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Abreviação:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   6960
         TabIndex        =   26
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Pessoa / Cliente / Razão Social:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Pessoas Cadastradas: 000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   195
      Left            =   7080
      TabIndex        =   21
      Top             =   120
      Width           =   2205
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   -120
      X2              =   9600
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Preencha o formulário"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1875
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BD As DAO.Database
Dim tb As DAO.Recordset

Private Sub cmdA_Click()
tb.MovePrevious
If tb.BOF = True Then
tb.MoveLast
End If
ExibirDados
cmdLimpar.Enabled = False
End Sub

Private Sub cmdAlterar_Click()

AbreCaixas
tb.Edit
cmdAlterar.Enabled = False
cmdGravar.Enabled = True
cmdBuscar.Enabled = False
cmdNovo.Enabled = False
cmdLimpar.Enabled = True
cmdExcluir.Enabled = False
cmda.Enabled = False
cmdp.Enabled = False
cmdSair.Caption = "&Cancelar"

End Sub

Private Sub cmdBusc_Click()
Dim ProcuraCodigo As String

    ProcuraCodigo = InputBox("Digite o código a ser consultado")
    tb.Seek "=", ProcuraCodigo
    If tb.NoMatch = True Then
    MsgBox ("Código não cadastrado!")
    tb.MovePrevious
    End If
    ExibirDados
End Sub

Private Sub cmdBuscar_Click()
frmBusca2.Show 1
End Sub

Private Sub cmdExcluir_Click()
If MsgBox("Confirma Exclusão desta Pessoa?", vbYesNo) = vbYes Then
tb.Delete
tb.MovePrevious
End If
cmdA_Click
Label3.Caption = "Pessoas Cadastradas: " & Format(tb.RecordCount, "000")
End Sub

Private Sub cmdGravar_Click()

grava
tb.Update
FechaCaixas
cmdGravar.Enabled = False
cmdAlterar.Enabled = True
cmdNovo.Enabled = True
cmdBuscar.Enabled = True
cmdExcluir.Enabled = True
cmdLimpar.Enabled = False
cmda.Enabled = True
cmdp.Enabled = True
cmdSair.Caption = "&Sair"
Label3.Caption = "Pessoas Cadastradas: " & Format(tb.RecordCount, "000")
MsgBox ("Dados cadastrados com sucesso...")

End Sub

Private Sub cmdLimpar_Click()
limpa
End Sub

Private Sub cmdNovo_Click()
tb.AddNew
Text1 = Format(tb.RecordCount, "000") + 4
AbreCaixas
limpa
Text2.SetFocus
cmdNovo.Enabled = False
cmdBuscar.Enabled = False
cmdAlterar.Enabled = False
cmdGravar.Enabled = True
cmdExcluir.Enabled = False
cmdLimpar.Enabled = False
cmda.Enabled = False
cmdp.Enabled = False
cmdSair.Caption = "&Cancelar"

End Sub

Private Sub cmdP_Click()
tb.MoveNext
If tb.EOF = True Then
tb.MovePrevious
End If
ExibirDados
cmdLimpar.Enabled = False
End Sub

Private Sub cmdsair_Click()
On Error Resume Next
If cmdSair.Caption = "&Cancelar" Then
cmdA_Click
cmdNovo.Enabled = True
cmdAlterar.Enabled = True
cmdGravar.Enabled = False
cmdExcluir.Enabled = True
cmdBuscar.Enabled = True
cmda.Enabled = True
cmdp.Enabled = True
FechaCaixas
cmdSair.Caption = "&Sair"
Else
If MsgBox("Quer sair do Cadastro de Pessoas?", vbYesNo, "Sair do Cadastro") = vbYes Then
    Unload frmClientes
    BD.Close
  Else
    Exit Sub
End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    SendKeys ("{TAB}")
    KeyAscii = 0
End If
End Sub

Private Sub Form_Load()
cmdNovo.Enabled = True
cmdSair.Enabled = True
FechaCaixas

Set BD = OpenDatabase(App.Path & "\pessoas.mdb")
Set tb = BD.OpenRecordset("Clientes", dbOpenTable)
tb.Index = "Codigo"

If tb.EOF = False Then
ExibirDados
cmdAlterar.Enabled = True
cmdExcluir.Enabled = True
cmdBuscar.Enabled = True
cmdLimpar.Enabled = False
cmda.Enabled = True
cmdp.Enabled = True
End If
Label3.Caption = "Pessoas Cadastradas: " & Format(tb.RecordCount, "000")
End Sub

Private Function ExibirDados()
On Error Resume Next
    Text1 = tb!codigo
    Text2 = tb!nome
    Text3 = tb!Opção
    Text4 = tb!Endereço
    Text5 = tb!Bairro
    Text6 = tb!Cidade
    Text7 = tb!Fone1
    Text8 = tb!observação
    Text9 = tb!uf
    Text10 = tb!cep
    Text11 = tb!Fone2

End Function

Private Function FechaCaixas()

Text1.Enabled = False
Text2.Enabled = False
Text2.BackColor = &H8000000A
Text3.Enabled = False
Text3.BackColor = &H8000000A
Text4.Enabled = False
Text4.BackColor = &H8000000A
Text5.Enabled = False
Text5.BackColor = &H8000000A
Text6.Enabled = False
Text6.BackColor = &H8000000A
Text7.Enabled = False
Text7.BackColor = &H8000000A
Text8.Enabled = False
Text8.BackColor = &H8000000A
Text9.Enabled = False
Text9.BackColor = &H8000000A
Text10.Enabled = False
Text10.BackColor = &H8000000A
Text11.Enabled = False
Text11.BackColor = &H8000000A

Label4.Enabled = False
Label5.Enabled = False
Label6.Enabled = False
Label7.Enabled = False
Label8.Enabled = False
Label9.Enabled = False
Label10.Enabled = False
Label11.Enabled = False
Label12.Enabled = False
Label13.Enabled = False

End Function

Private Function AbreCaixas()

Text2.Enabled = True
Text2.BackColor = &H80000005
Text3.Enabled = True
Text3.BackColor = &H80000005
Text4.Enabled = True
Text4.BackColor = &H80000005
Text5.Enabled = True
Text5.BackColor = &H80000005
Text6.Enabled = True
Text6.BackColor = &H80000005
Text7.Enabled = True
Text7.BackColor = &H80000005
Text8.Enabled = True
Text8.BackColor = &H80000005
Text9.Enabled = True
Text9.BackColor = &H80000005
Text10.Enabled = True
Text10.BackColor = &H80000005
Text11.Enabled = True
Text11.BackColor = &H80000005

Label4.Enabled = True
Label5.Enabled = True
Label6.Enabled = True
Label7.Enabled = True
Label8.Enabled = True
Label9.Enabled = True
Label10.Enabled = True
Label11.Enabled = True
Label12.Enabled = True
Label13.Enabled = True

End Function

Private Function limpa()

Text2 = Empty
Text3 = Empty
Text4 = Empty
Text5 = Empty
Text6 = Empty
Text7 = Empty
Text8 = Empty
Text9 = Empty
Text10 = Empty
Text11 = Empty

End Function

Private Function grava()

tb("codigo") = Text1
tb("nome") = Text2
tb("opção") = Text3
tb("endereço") = Text4
tb("bairro") = Text5
tb("cidade") = Text6
tb("fone1") = Text7
tb("observação") = Text8
tb("uf") = Text9
tb("cep") = Text10
tb("fone2") = Text11

End Function

Private Sub Text1_Change()
Text1 = Format(Text1.text, "000")
End Sub

Private Sub Text2_Change()
If Text2.text = "" Then
cmdLimpar.Enabled = False
Else
cmdLimpar.Enabled = True
End If
End Sub

Private Sub Text2_GotFocus()
Text2.BackColor = &HE0E0E0
Text2.Font.Name = "Tahoma"
Text2.Font.Bold = True
End Sub

Private Sub Text2_LostFocus()
Text2 = StrConv(Text2.text, vbUpperCase)
Text2.BackColor = &H80000005
Text2.Font.Name = "Ms Sans Serif"
Text2.Font.Bold = False
End Sub

Private Sub Text3_GotFocus()
Text3.BackColor = &HE0E0E0
Text3.Font.Name = "Tahoma"
Text3.Font.Bold = True
End Sub

Private Sub Text3_LostFocus()
Text3 = StrConv(Text3.text, vbUpperCase)
Text3.BackColor = &H80000005
Text3.Font.Name = "Ms Sans Serif"
Text3.Font.Bold = False
End Sub

Private Sub Text4_GotFocus()
Text4.BackColor = &HE0E0E0
Text4.Font.Name = "Tahoma"
Text4.Font.Bold = True
End Sub

Private Sub Text4_LostFocus()
Text4 = StrConv(Text4.text, vbUpperCase)
Text4.BackColor = &H80000005
Text4.Font.Name = "Ms Sans Serif"
Text4.Font.Bold = False
End Sub

Private Sub Text5_GotFocus()
Text5.BackColor = &HE0E0E0
Text5.Font.Name = "Tahoma"
Text5.Font.Bold = True
End Sub

Private Sub Text5_LostFocus()
Text5 = StrConv(Text5.text, vbUpperCase)
Text5.BackColor = &H80000005
Text5.Font.Name = "Ms Sans Serif"
Text5.Font.Bold = False
End Sub

Private Sub Text6_GotFocus()
Text6.BackColor = &HE0E0E0
Text6.Font.Name = "Tahoma"
Text6.Font.Bold = True
End Sub

Private Sub Text6_LostFocus()
Text6 = StrConv(Text6.text, vbUpperCase)
Text6.BackColor = &H80000005
Text6.Font.Name = "Ms Sans Serif"
Text6.Font.Bold = False
End Sub

Private Sub Text7_GotFocus()
Text7.BackColor = &HE0E0E0
Text7.Font.Name = "Tahoma"
Text7.Font.Bold = True
End Sub

Private Sub Text7_LostFocus()
Text7 = StrConv(Text7.text, vbUpperCase)
Text7.BackColor = &H80000005
Text7.Font.Name = "Ms Sans Serif"
Text7.Font.Bold = False
End Sub

Private Sub Text8_GotFocus()
Text8.BackColor = &HE0E0E0
Text8.Font.Name = "Tahoma"
Text8.Font.Bold = True
End Sub

Private Sub Text8_LostFocus()
Text8 = StrConv(Text8.text, vbUpperCase)
Text8.BackColor = &H80000005
Text8.Font.Name = "Ms Sans Serif"
Text8.Font.Bold = False
End Sub

Private Sub Text9_GotFocus()
Text9.BackColor = &HE0E0E0
Text9.Font.Name = "Tahoma"
Text9.Font.Bold = True
End Sub

Private Sub Text9_LostFocus()
Text9 = StrConv(Text9.text, vbUpperCase)
Text9.BackColor = &H80000005
Text9.Font.Name = "Ms Sans Serif"
Text9.Font.Bold = False
End Sub

Private Sub Text10_GotFocus()
Text10.BackColor = &HE0E0E0
Text10.Font.Name = "Tahoma"
Text10.Font.Bold = True
End Sub

Private Sub Text10_LostFocus()
Text10 = StrConv(Text10.text, vbUpperCase)
Text10.BackColor = &H80000005
Text10.Font.Name = "Ms Sans Serif"
Text10.Font.Bold = False
End Sub

Private Sub Text11_GotFocus()
Text11.BackColor = &HE0E0E0
Text11.Font.Name = "Tahoma"
Text11.Font.Bold = True
End Sub

Private Sub Text11_LostFocus()
Text11 = StrConv(Text11.text, vbUpperCase)
Text11.BackColor = &H80000005
Text11.Font.Name = "Ms Sans Serif"
Text11.Font.Bold = False
End Sub
