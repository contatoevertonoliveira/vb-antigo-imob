VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFaturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gerar Recibos"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      Height          =   675
      Left            =   360
      Picture         =   "frmFaturas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Gerar Recibos"
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Height          =   675
      Left            =   5400
      Picture         =   "frmFaturas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Sair do Gerador"
      Top             =   4680
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   4440
      Width           =   6375
      Begin VB.CommandButton Command5 
         Enabled         =   0   'False
         Height          =   675
         Left            =   4080
         Picture         =   "frmFaturas.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Ver Recibo"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Height          =   675
         Left            =   2760
         Picture         =   "frmFaturas.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Cadastrar Vencimentos"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Enabled         =   0   'False
         Height          =   675
         Left            =   1440
         Picture         =   "frmFaturas.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprimir Recibo"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtvalor 
      Alignment       =   2  'Center
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5280
      TabIndex        =   3
      ToolTipText     =   "Valor do Aluguél"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtNumeroParcelas 
      Alignment       =   2  'Center
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   3720
      TabIndex        =   2
      ToolTipText     =   "Nº de Recibos"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtNumeroTitulo 
      Alignment       =   2  'Center
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Tempo  de Contrato"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtData 
      Alignment       =   2  'Center
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   1
      ToolTipText     =   "Início do Contrato"
      Top             =   1800
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstFaturas 
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4048
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nº Contrato"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Parcelas"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Vencimento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Valor"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.Line Line4 
      X1              =   1440
      X2              =   1920
      Y1              =   840
      Y2              =   1080
   End
   Begin VB.Line Line3 
      X1              =   1440
      X2              =   1920
      Y1              =   840
      Y2              =   480
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Locatário:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2040
      TabIndex        =   20
      Top             =   960
      Width           =   870
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Locador:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2040
      TabIndex        =   19
      Top             =   360
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   660
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   6720
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   6720
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3120
      TabIndex        =   17
      Top             =   360
      Width           =   465
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3120
      TabIndex        =   16
      Top             =   960
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   840
      TabIndex        =   10
      Top             =   720
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Valor Aluguél:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5280
      TabIndex        =   9
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nº Parcelas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3720
      TabIndex        =   8
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicio:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2040
      TabIndex        =   7
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label txtTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Tempo / Contrato:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "frmFaturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bd As DAO.Database
Dim Tbl As DAO.Recordset
Dim slabel12 As String
Dim slabel13 As String
Dim slabel14 As String
Dim slabel15 As String
Dim slabel16 As String
Dim slabel17 As String
Dim slabel18 As String
Dim i As Long, QtdeCol As Long, QtdeTxt, ii As Integer, achou As Boolean




Private Sub Command1_Click()

If Len(txtNumeroTitulo) < 1 Then
MsgBox ("Preencha o prazo do contrato!")
txtNumeroTitulo.SetFocus
End If

If Len(txtData) < 1 Then
MsgBox ("Preencha a data inicial do contrato!")
txtData.SetFocus
End If

If Len(txtvalor) < 1 Then
MsgBox ("Preencha o valor do contrato!")
txtvalor.SetFocus
End If

If Len(txtNumeroParcelas) < 1 Then
MsgBox ("Preencha o nº de parcelas!")
txtNumeroParcelas.SetFocus
End If

lstFaturas.ListItems.Clear
If Label4.Caption = "" Or txtvalor.Text = "" Or txtNumeroParcelas.Text = "" Or txtNumeroTitulo.Text = "" Then Exit Sub

'================================================================
'  Gera fatura para ano com 2 e 4 dígitos
'================================================================

Dim DataFatura, DataParcela As Date
Dim ValorFatura As Currency
Dim DD, MM, AA, Parcela, NumeroParcela As Integer

'======================================

DataFatura = txtData.Text
ValorFatura = txtvalor.Text
NumeroParcela = txtNumeroParcelas.Text
Parcela = 0

'======================================

DD = Mid$(DataFatura, 1, 2)
MM = Mid$(DataFatura, 4, 2)
AA = Mid$(DataFatura, 7, 4)

'======================================

Do Until Parcela >= NumeroParcela

    DD = Format(CCur(DD), "##")
    If MM <= 12 Then MM = CCur(MM) + 1
    If CCur(MM) > 12 Then
        MM = CCur(1)
        AA = CCur(AA) + 1
    End If
    
    Parcela = Parcela + 1

    DataParcela = CDate(DD & "/" & MM & "/" & AA)

'=======================================

Dim lst As ListItem
Set lst = lstFaturas.ListItems.Add(, "000" + Str(lstFaturas.ListItems.Count + 1), Str(lstFaturas.ListItems.Count + 1))
    lst.ListSubItems.Add Text:=Label4.Caption
    lst.ListSubItems.Add Text:=lstFaturas.ListItems.Count & "/" & txtNumeroTitulo.Text
    lst.ListSubItems.Add Text:=CDate(DataParcela)
    lst.ListSubItems.Add Text:=Format(CCur(ValorFatura), "R$#,##0.00;(R$#,##0.00)")
Loop

Command1.Enabled = False
Command4.Enabled = True
End Sub

Private Sub Command2_Click()

If MsgBox("Quer sair do Gera Recibos?", vbYesNo, "Sair do Gerador") = vbYes Then
    Unload Me
    FechaLabels
    frmRecibos.Text2.Visible = False
    frmRecibos.Text1 = Empty
    frmRecibos.Command3.Picture = LoadPicture("C:\Programa Imobiliária\Ícones\ARW10NE.ICO")
    frmRecibos.Command3.Caption = "S&air"
    frmRecibos.Command1.Enabled = False
    frmRecibos.Command4.Visible = False
    frmRecibos.Label3.Visible = False
Else
    Exit Sub
End If
End Sub

Private Sub Command4_Click()

Tbl.AddNew
If Label4 <> "" Then Tbl("codigo") = Label4
If Label5 <> "" Then Tbl("Locatario") = Label5
If Label6 <> "" Then Tbl("Locador") = Label6
If frmRecibos.Label26 <> "" Then Tbl("Iptu") = frmRecibos.Label26
If frmRecibos.Label27 <> "" Then Tbl("Multa") = frmRecibos.Label27
If txtNumeroParcelas <> "" Then Tbl("Nrecibos") = txtNumeroParcelas
If txtvalor <> "" Then Tbl("valor") = txtvalor
If txtData <> "" Then Tbl("Dtinicial") = txtData
If frmRecibos.Label29 <> "" Then Tbl("Dtfinal") = frmRecibos.Label29
Tbl.Update

MsgBox ("Recibos cadastrados com sucesso!")
Command4.Enabled = False
End Sub

Private Sub Command5_Click()
Dim dia As String
Dim mes As String
Dim ano As String
Dim Primeiro As String
Dim Segundo As String

If lstFaturas.SelectedItem.Text <> "" Then
    frmRecibos.Label32.Visible = True
    frmRecibos.Label32 = DateAdd("m", lstFaturas.SelectedItem.Text, CDate(frmRecibos.Label23))
        
        frmRecibos.Label18 = CDate(frmRecibos.Label32)
        dia = Day(frmRecibos.Label32)
        mes = Month(frmRecibos.Label32) - 1
        ano = Year(frmRecibos.Label32)
            If mes < 1 Then
                mes = 12
                ano = ano - 1
            End If
        
        frmRecibos.Label18.Visible = True
        frmRecibos.Label18 = dia & "/" & mes & "/" & ano & "       À"
        Primeiro = Format(frmRecibos.Label25, "Currency")
        Segundo = Format(frmRecibos.Label26, "Currency")
        
        frmRecibos.Label28.Visible = True
        frmRecibos.Label28 = Primeiro + Segundo
        
    frmRecibos.Label24.Caption = "Aluguél: " & Format(lstFaturas.SelectedItem.Text, "00") & "/" & txtNumeroParcelas
End If
Me.Hide
End Sub

Private Sub Form_Load()

Set Bd = OpenDatabase(App.Path & "\dados\bdimobiliaria.MDB")
Set Tbl = Bd.OpenRecordset("Vencimentos", dbOpenTable)

End Sub

Private Sub lstFaturas_Click()
Command5.Enabled = True
End Sub

Private Sub lstFaturas_DblClick()
Dim dia As String
Dim mes As String
Dim ano As String

If lstFaturas.SelectedItem.Text <> "" Then
    frmRecibos.Label32.Visible = True
    frmRecibos.Label32 = DateAdd("m", lstFaturas.SelectedItem.Text, CDate(frmRecibos.Label23))
        
        frmRecibos.Label18 = CDate(frmRecibos.Label32)
        dia = Day(frmRecibos.Label32)
        mes = Month(frmRecibos.Label32) - 1
        ano = Year(frmRecibos.Label32)
            If mes < 1 Then
                mes = 12
                ano = ano - 1
            End If
        
        frmRecibos.Label18.Visible = True
        frmRecibos.Label18 = dia & "/" & mes & "/" & ano & "       À"
        
    frmRecibos.Label24.Caption = "Aluguél: " & Format(lstFaturas.SelectedItem.Text, "00") & "/" & txtNumeroParcelas
End If
Me.Hide
End Sub

Private Sub txtData_Change()
txtData = Format$(txtData, "dd/mm/yyyy")
End Sub

Private Sub txtData_GotFocus()
txtData.BackColor = &HFFFF&
End Sub

Private Sub txtData_LostFocus()

txtData.BackColor = &H80000005

If Not IsDate(txtData.Text) Then
    MsgBox "Data Inválida"
    txtData.SetFocus
End If

End Sub

Private Sub txtNumeroParcelas_GotFocus()
txtNumeroParcelas.BackColor = &HFFFF&
End Sub

Private Sub txtNumeroParcelas_LostFocus()
txtNumeroParcelas.BackColor = &H80000005
End Sub

Private Sub txtNumeroTitulo_Change()
txtNumeroParcelas = txtNumeroTitulo
End Sub

Private Sub txtNumeroTitulo_GotFocus()
txtNumeroTitulo.BackColor = &HFFFF&
End Sub

Private Sub txtNumeroTitulo_LostFocus()
txtNumeroTitulo.BackColor = &H80000005
End Sub

Private Sub txtvalor_Change()
If Len(txtvalor) > 0 Then
Command1.Enabled = True
End If
End Sub

Private Sub txtvalor_GotFocus()
txtvalor.BackColor = &HFFFF&
End Sub

Private Sub txtvalor_LostFocus()
txtvalor.BackColor = &H80000005
End Sub

Private Function FechaLabels()
Dim LabelObjeto As Object

For Each LabelObjeto In frmRecibos.Controls
If TypeOf LabelObjeto Is Label Then
    LabelObjeto.Visible = False
End If
Next LabelObjeto

Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
End Function
