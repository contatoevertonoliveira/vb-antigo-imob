VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGrids 
   Caption         =   "Consulta de Clientes Cadatrados"
   ClientHeight    =   8100
   ClientLeft      =   90
   ClientTop       =   1515
   ClientWidth     =   11865
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtBusca 
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdParar 
      Caption         =   "Command2"
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdBusca 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   7560
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MsFlexGrid1 
      Height          =   3615
      Left            =   480
      TabIndex        =   0
      Top             =   2280
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6376
      _Version        =   393216
   End
   Begin VB.Label lblMsflexgrid 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   720
      TabIndex        =   8
      Top             =   2040
      Width           =   465
   End
   Begin VB.Label lblBusca 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   720
      TabIndex        =   7
      Top             =   1800
      Width           =   465
   End
End
Attribute VB_Name = "frmGrids"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Integer
Dim strTextoBusca As String
Dim strBusca As String
Dim FNome As Boolean
Dim FSetor As Boolean
Dim FContato As Boolean


Private Sub cmdBusca_Click()
On Error GoTo ErrorHandler ' tratamento de erros

If txtBusca = "" Or IsNumeric(txtBusca.Text) Then
   MsgBox "Informe um texto alfanumérico válido !", vbCritical, "Erro"
   txtBusca.Text = ""
   txtBusca.SetFocus
   Exit Sub
End If

If cmdParar.Caption = Cap4 Then
  cmdParar.Caption = Cap41
  Screen.MousePointer = vbHourglass

  strTextoBusca = Trim(txtBusca.Text)
  rstObj.Close
  Set rstObj = Nothing
  If FNome = True Then
     strBusca = "SELECT * FROM Employees where UserName Like '" & strTextoBusca & "%' Order By UserName"
     FSetor = False
     FContato = False
  End If
  If FSetor = True Then
     strBusca = "SELECT * FROM Employees where Department Like '" & strTextoBusca & "%' Order By Department"
     FNome = False
     FContato = False
  End If
  If FContato = True Then
      strBusca = "SELECT * FROM Employees where ContactPerson Like '" & strTextoBusca & "%' Order By ContactPerson"

      FNome = False
      FSetor = False
  End If

 Call AbrirRecordSetAccess(strBusca)
 MsFlexGrid1.Refresh
 Call PreencherMSFlexGrid
 DataGrid1.Refresh
 Call PreencherDataGrid

 Animation1.Visible = False
 Screen.MousePointer = vbDefault
 cmdParar.Caption = Cap4
End If
Exit Sub
ErrorHandler:  'inicio do tratamento de erros
  MsgBox "Erro  No. :" & Err.Number & vbCr & " Descrição :" & Err.Description
  Animation1.Visible = False
  Screen.MousePointer = vbDefault
  Resume ' retorna a execução para a mesma linha onde ocorreu o erro
  cnxnObj.Close
  Set cnxnObj = Nothing


End Sub

Private Sub Form_Load()

frmGrids.Caption = Cap1
frmGrids.WindowState = 0
  lblBusca.Caption = Cap2
  cmdBusca.Caption = Cap3
  cmdParar.Caption = Cap4
  lblMsflexgrid.Caption = Cap6
  Option1.Caption = Cap8
  Option2.Caption = Cap9
  Option3.Caption = Cap10
  Option1.Value = True

Call AbrirBDAccess
Call AbrirRecordSetAccess("SELECT * FROM Employees Order By UserName")
Call PreencherMSFlexGrid

End Sub


Sub PreencherMSFlexGrid()
 MsFlexGrid1.Cols = 4
 MsFlexGrid1.ColWidth(0) = 500
 MsFlexGrid1.TextMatrix(0, 0) = "Sr.No"
 For I = 0 To rstObj.Fields.Count - 1
   MsFlexGrid1.ColAlignment(I) = vbCenter
   MsFlexGrid1.ColWidth(I + 1) = 1500
   MsFlexGrid1.TextMatrix(0, I + 1) = rstObj.Fields(I).Name
 Next
 MsFlexGrid1.Rows = rstObj.RecordCount + 1
 I = 1
 Do While Not rstObj.EOF
    MsFlexGrid1.TextMatrix(I, 0) = I
    MsFlexGrid1.TextMatrix(I, 1) = rstObj(0) 'username
    MsFlexGrid1.TextMatrix(I, 2) = rstObj(1) 'Department
    MsFlexGrid1.TextMatrix(I, 3) = rstObj(2) 'ContactPerson
    I = I + 1
    rstObj.MoveNext
 Loop
End Sub

