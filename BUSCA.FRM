VERSION 5.00
Begin VB.Form BUSCA 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Localizador de Recibos"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6030
   Icon            =   "BUSCA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
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
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   4935
   End
   Begin VB.TextBox Busca 
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   5775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Efetuar Busca"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   2745
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Cancelar Busca"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   2790
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   1215
   End
End
Attribute VB_Name = "BUSCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()

If Combo1.Text = "Locador" Then
    frmAluguel.Data1.RecordSource = "SELECT * FROM vencimentos WHERE Locador = '" & Busca.Text & "'"
    frmAluguel.Data1.Refresh
    If frmAluguel.Data1.Recordset.EOF = True Then
        frmAluguel.Data1.RecordSource = "SELECT * FROM Vencimentos"
        frmAluguel.Data1.Refresh
        MsgBox ("Nenhum registro foi encontrado"), vbCritical, "Registro não encontrado"
        Busca.Text = ""
    Else
    Unload Me
    End If
    
ElseIf Combo1.Text = "Locador" Then
    frmAluguel.Data1.RecordSource = "SELECT * FROM vencimentos WHERE Locador = '" & Busca.Text & "'"
    frmAluguel.Data1.Refresh
    If frmAluguel.Data1.Recordset.EOF = True Then
        frmAluguel.Data1.RecordSource = "SELECT * FROM Vencimentos"
        frmAluguel.Data1.Refresh
        MsgBox ("Nenhum registro foi encontrado"), vbCritical, "Registro não encontrado"
        Busca.Text = ""
    Else
    Unload Me
    End If
    
ElseIf Combo1.Text = "Locatário" Then
    frmAluguel.Data1.RecordSource = "SELECT * FROM Vencimentos WHERE Locatario = '" & Busca.Text & "'"
    frmAluguel.Data1.Refresh
    If frmAluguel.Data1.Recordset.EOF = True Then
        frmAluguel.Data1.RecordSource = "SELECT * FROM Vencimentos"
        frmAluguel.Data1.Refresh
        MsgBox ("Nenhum registro foi encontrado"), vbCritical, "Registro não encontrado"
        Busca.Text = ""
    Else
    Unload Me
    End If
  
ElseIf Combo1.Text = "Data de Lançamento" Then
    frmAluguel.Data1.RecordSource = "SELECT * FROM a_pagar WHERE Descrição = #" & Busca.Text & "#"
    frmAluguel.Data1.Refresh
    If frmAluguel.Data1.Recordset.EOF = True Then
       frmAluguel.Data1.RecordSource = "SELECT * FROM a_pagar"
        frmAluguel.Data1.Refresh
        MsgBox ("Nenhum registro foi encontrado"), vbCritical, "Registro não encontrado"
        Busca.Text = ""
    Else
    Unload Me
    End If
    
ElseIf Combo1.Text = "Data de Pagamento" Then
    frmAluguel.Data1.RecordSource = "SELECT * FROM a_pagar WHERE Descrição = #" & Busca.Text & "#"
    frmAluguel.Data1.Refresh
    If frmAluguel.Data1.Recordset.EOF = True Then
        frmAluguel.Data1.RecordSource = "SELECT * FROM a_pagar"
        frmAluguel.Data1.Refresh
        MsgBox ("Nenhum registro foi encontrado"), vbCritical, "Registro não encontrado"
        Busca.Text = ""
    Else
    Unload Me
    End If

End If
End Sub

Private Sub Form_Load()

With Combo1
    .AddItem "Selecione uma opção:"
    .AddItem "Código"
    .AddItem "Locador"
    .AddItem "Locatário"
End With

Combo1.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub


Private Function VerificaMes()

If Busca.Text = "Janeiro" Then
   Busca = Month(1)
   ElseIf Busca.Text = "Fevereiro" Then
   Busca = Month(2)
   ElseIf Busca.Text = "Março" Then
   Busca = Month(3)
   ElseIf Busca.Text = "Abril" Then
   Busca = Month(4)
   ElseIf Busca.Text = "Maio" Then
   Busca = Month(5)
   ElseIf Busca.Text = "Junho" Then
   Busca = Month(6)
   ElseIf Busca.Text = "Julho" Then
   Busca = Month(7)
   ElseIf Busca.Text = "Agosto" Then
   Busca = Month(8)
   ElseIf Busca.Text = "Setembro" Then
   Busca = Month(9)
   ElseIf Busca.Text = "Outubro" Then
   Busca = Month(10)
   ElseIf Busca.Text = "Novembro" Then
   Busca = Month(11)
   ElseIf Busca.Text = "Dezembro" Then
   Busca = Month(12)
End If

End Function
