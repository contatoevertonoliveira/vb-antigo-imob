VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmVencimentos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Super Imob - Gerador de Recibos"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programa Imobiliária\Dados\Bdimobiliaria.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Vencimentos"
      Top             =   4920
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Frame Frame3 
      Caption         =   "Recibos Gerados:"
      Height          =   1935
      Left            =   120
      TabIndex        =   25
      Top             =   3120
      Width           =   8295
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmVencimentos.frx":0000
         Height          =   1335
         Left            =   120
         OleObjectBlob   =   "frmVencimentos.frx":0014
         TabIndex        =   28
         Top             =   360
         Width           =   8055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados dos Recibos:"
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   1920
      Width           =   8295
      Begin VB.CommandButton Command3 
         Caption         =   "&Organizar Recibos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5880
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ger&ar + Recibos"
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
         Height          =   735
         Left            =   7080
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ger&ar Recibos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4800
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2160
         TabIndex        =   24
         Top             =   720
         Width           =   345
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Data Final:"
         Height          =   195
         Left            =   1200
         TabIndex        =   23
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2160
         TabIndex        =   22
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial:"
         Height          =   195
         Left            =   1200
         TabIndex        =   21
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4080
         TabIndex        =   20
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Aluguél R$:"
         Height          =   195
         Left            =   3120
         TabIndex        =   19
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8295
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   6240
         TabIndex        =   6
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   7200
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   7200
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Os vencimentos já foram gerados !"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   2760
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         Height          =   195
         Left            =   5640
         TabIndex        =   15
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Aluguél R$:"
         Height          =   195
         Left            =   6240
         TabIndex        =   14
         Top             =   720
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Iptu:"
         Height          =   195
         Left            =   6720
         TabIndex        =   13
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Locatário:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Locador:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   630
      End
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   8280
      TabIndex        =   29
      Top             =   5040
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "frmVencimentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CN As DAO.Database
Dim Rs As DAO.Recordset

Private Sub Command1_Click()
Dim mes As Integer
Dim ANO As Integer
Dim Data As String
Dim Valor As String

If Text8.Text = frmRecibos.Label30 Then

mes = Format(frmRecibos.Label23, "mm")
ANO = Format(frmRecibos.Label23, "yy")
Valor = Val("0")

For i = 1 To Val(Text8.Text)
    mes = mes + 1
    Valor = Valor + 1
    If mes > 12 Then
       mes = 1
       ANO = ANO + 1
   End If
    
  dia1 = Format(frmRecibos.Label23, "dd")
  dia = Verifica_dia(dia1, mes)
  Data = dia & "/" & Format(mes, "0") & "/" & Format(ANO, "0000")
  
  Data2.Recordset.AddNew
  Data2.Recordset.Fields(1) = CLng(Text7)
  Data2.Recordset.Fields(2) = CDate(Data)
  Data2.Recordset.Fields(3) = Text1
  Data2.Recordset.Fields(4) = Text2
  Data2.Recordset.Fields(5) = Text4
  Data2.Recordset.Fields(6) = Text5
  Data2.Recordset.Fields(7) = Label12
  Data2.Recordset.Fields(8) = Label14
  Data2.Recordset.Fields(11) = Val(Valor)
  Data2.Recordset.Update
Next
MsgBox ("Recibos Gerados com sucesso!")

Else
    MsgBox ("O número que você digitou não é igual ao prazo do contrato, por favor corrigir!")
    Text8 = ""
    Text8.SetFocus
End If

Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
End Sub

Private Sub Data1_Reposition()
If Text7.Text <> "" Then
Data2.RecordSource = "Select * from Vencimentos where Codigo= " & CLng(Text7.Text)
Data2.Refresh
End If
End Sub

Private Sub Command2_Click()
frmRec.Text4 = Label10
frmRec.Show 1
End Sub

Private Sub Command3_Click()
If Text1.Text = "" Then
    Data2.RecordSource = "SELECT * FROM VENCIMENTOS"
    Data2.Refresh
    Exit Sub
End If

Data2.RecordSource = "SELECT * FROM VENCIMENTOS WHERE Locador Like '" & Text1.Text & "*'"
Data2.Refresh
End Sub

Private Sub DBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
On Error GoTo erro

If ColIndex >= 0 And ColIndex <= 5 Then
    Cancel = True
    MsgBox "Não pode ser alterado o conteudo desta célula.", vbCritical, "Aviso!"
    Exit Sub
End If

Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Aviso": Exit Sub
End Sub

Private Sub DBGrid1_DblClick()
On Error Resume Next
Dim dia As String
Dim mes As String
Dim ANO As String
Dim resultado As String
Dim Prazo As String

Prazo = frmRecibos.Label30.Caption
frmRecibos.Label32.Visible = True
If Data2.Recordset.Fields(2) < Format(Now, "dd") - 3 Then
    frmRecibos.Label31 = Data2.Recordset.Fields(1)
    frmRecibos.Label32 = Data2.Recordset.Fields(2)
    frmRecibos.Label15 = Data2.Recordset.Fields(3)
    frmRecibos.Label16 = Data2.Recordset.Fields(4)
    frmRecibos.Label26 = Data2.Recordset.Fields(5)
    frmRecibos.Label25 = Data2.Recordset.Fields(6)
    frmRecibos.Label24 = "Aluguél: " & Format(Data2.Recordset.Fields(11), "00") & "/" & Val(Prazo)
    dia = Day(frmRecibos.Label32)
    mes = Month(frmRecibos.Label32) - 1
    ANO = Year(frmRecibos.Label32)
    If mes < 1 Then
       mes = 12
       ANO = ANO - 1
    End If
    Unload Me
    frmRecibos.Label18.Visible = True
    frmRecibos.Label18 = dia & "/" & mes & "/" & ANO
    frmRecibos.Command5.Enabled = True
    frmRecibos.Command2.Enabled = True
    frmRecibos.Command3.Enabled = True
Else
    frmRecibos.Label31 = Data2.Recordset.Fields(1)
    frmRecibos.Label32 = CDate(Data2.Recordset.Fields(2))
    frmRecibos.Label15 = Data2.Recordset.Fields(3)
    frmRecibos.Label16 = Data2.Recordset.Fields(4)
    frmRecibos.Label26 = Data2.Recordset.Fields(5)
    frmRecibos.Label25 = Data2.Recordset.Fields(6)
    frmRecibos.Label24 = "Aluguél: " & Format(Data2.Recordset.Fields(11), "00") & "/" & Val(Prazo)
    dia = Day(frmRecibos.Label32)
    mes = Month(frmRecibos.Label32) - 1
    ANO = Year(frmRecibos.Label32)
    If mes < 1 Then
       mes = 12
       ANO = ANO - 1
    End If
    Unload Me
    frmRecibos.Label18.Visible = True
    frmRecibos.Label18 = dia & "/" & mes & "/" & ANO
    frmRecibos.Command5.Enabled = True
    frmRecibos.Command2.Enabled = True
    frmRecibos.Command3.Enabled = True
End If
frmRecibos.Command2.Enabled = True
frmRecibos.Command9.Enabled = True
End Sub

Private Sub Form_Load()

Data2.DatabaseName = App.Path & "\Dados\Bdimobiliaria.MDB"
Data2.RecordSource = "Vencimentos"

End Sub


Public Function Verifica_dia(dia, mes)
Dim diasDoMes As Variant

dia = Val(dia)

diasDoMes = Array(31, 28, 30, 30, 31, 30, 31, 30, 30, 31, 30, 31)

If dia = 31 Then
Verifica_dia = diasDoMes(mes - 1)
Else
Verifica_dia = dia
End If

End Function

Private Sub Form_Unload(Cancel As Integer)
Label16.Caption = "0"
Data2.Refresh
End Sub

Private Sub Label16_Change()
If Label16 = "1" Then
    Label15.Visible = True
    Text1.BackColor = &HFFFF&
    Text1.Enabled = False
    Text2.BackColor = &HFFFF&
    Text2.Enabled = False
    Text3.BackColor = &HFFFF&
    Text3.Enabled = False
    Text4.BackColor = &HFFFF&
    Text4.Enabled = False
    Text5.BackColor = &HFFFF&
    Text5.Enabled = False
    Text6.BackColor = &HFFFF&
    Text6.Enabled = False
    Text7.BackColor = &HFFFF&
    Text7.Enabled = False
    Command1.Enabled = False
    Command2.Enabled = True
ElseIf Label16 = "0" Then
    Label15.Visible = False
    Text1.BackColor = &H80000005
    Text1.Enabled = True
    Text2.BackColor = &H80000005
    Text2.Enabled = True
    Text3.BackColor = &H80000005
    Text3.Enabled = True
    Text4.BackColor = &H80000005
    Text4.Enabled = True
    Text5.BackColor = &H80000005
    Text5.Enabled = True
    Text6.BackColor = &H80000005
    Text6.Enabled = True
    Text7.BackColor = &H80000005
    Text7.Enabled = True
    Command1.Enabled = True
    Command2.Enabled = False
End If
End Sub

Private Sub Text5_Change()
Label10 = Text5
End Sub

Private Function FechaCaixas()
Dim P_objeto As Object

For Each P_objeto In Me.Controls
    If TypeOf P_objeto Is TextBox Then
    P_objeto.Enabled = False
    P_objeto.BackColor = &H80000013
End If
Next P_objeto

End Function
