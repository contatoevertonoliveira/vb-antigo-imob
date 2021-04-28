VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRecibos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Super Imob - Emissão de Recibos"
   ClientHeight    =   7290
   ClientLeft      =   1650
   ClientTop       =   1215
   ClientWidth     =   8820
   Icon            =   "frmRecibos2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command8 
      Caption         =   "Pr&estação"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6600
      TabIndex        =   41
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "L&ocação"
      Height          =   495
      Left            =   3480
      TabIndex        =   40
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Re&cibo Universal"
      Height          =   495
      Left            =   240
      TabIndex        =   39
      Top             =   120
      Width           =   1935
   End
   Begin VB.Data dtaContrato 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programa Imobiliária\Dados\Bdimobiliaria.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Contrato"
      Top             =   6960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   8295
      Begin VB.CommandButton Command9 
         Caption         =   "Calcul&ar Recibo"
         Enabled         =   0   'False
         Height          =   855
         Left            =   6600
         Picture         =   "frmRecibos2.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Edit&ar Recibo"
         Enabled         =   0   'False
         Height          =   855
         Left            =   6600
         Picture         =   "frmRecibos2.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "L&impar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3720
         TabIndex        =   30
         Top             =   3720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Sa&ir"
         Height          =   615
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Imprimir Recibo"
         Enabled         =   0   'False
         Height          =   855
         Left            =   6600
         Picture         =   "frmRecibos2.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ger&ar Recibos"
         Enabled         =   0   'False
         Height          =   855
         Left            =   6600
         Picture         =   "frmRecibos2.frx":1330
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2775
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   960
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Label30"
         DataField       =   "Prazo"
         DataSource      =   "dtaContrato"
         Height          =   195
         Left            =   5880
         TabIndex        =   45
         Top             =   4080
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "Label25"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3720
         TabIndex        =   42
         Top             =   2520
         Width           =   810
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Label32"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2640
         TabIndex        =   37
         Top             =   2520
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Label31"
         DataField       =   "ID"
         DataSource      =   "dtaContrato"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1080
         TabIndex        =   36
         Top             =   960
         Width           =   570
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Label29"
         DataField       =   "Final"
         DataSource      =   "dtaContrato"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2280
         TabIndex        =   35
         Top             =   4080
         Width           =   570
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Label23"
         DataField       =   "Inicio"
         DataSource      =   "dtaContrato"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2280
         TabIndex        =   34
         Top             =   3840
         Width           =   570
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Label22"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   960
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "Label28"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3720
         TabIndex        =   28
         Top             =   3240
         Width           =   810
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "0,00"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3720
         TabIndex        =   27
         Top             =   3000
         Width           =   810
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Label26"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3720
         TabIndex        =   26
         Top             =   2760
         Width           =   810
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Label24"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3600
         TabIndex        =   22
         Top             =   4080
         Width           =   570
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Label21"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1320
         TabIndex        =   21
         Top             =   3240
         Width           =   570
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Label20"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1320
         TabIndex        =   20
         Top             =   3000
         Width           =   570
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Label19"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1320
         TabIndex        =   19
         Top             =   2760
         Width           =   570
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Label18"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1320
         TabIndex        =   18
         Top             =   2520
         Width           =   570
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Label17"
         DataField       =   "ImovelLocado"
         DataSource      =   "dtaContrato"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1680
         TabIndex        =   17
         Top             =   1800
         Width           =   570
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Label16"
         DataField       =   "aLocatario"
         DataSource      =   "dtaContrato"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1680
         TabIndex        =   16
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Label15"
         DataField       =   "Locador"
         DataSource      =   "dtaContrato"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1680
         TabIndex        =   15
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Label14"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   4080
         Width           =   570
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Label13"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   3840
         Width           =   570
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Label12"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   3480
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Label11"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   3240
         Width           =   570
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Label10"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   3000
         Width           =   570
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Label9"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   2760
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Label8"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Label7"
         Height          =   195
         Left            =   1680
         TabIndex        =   7
         Top             =   2160
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Label6"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Label5"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Label4"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3240
         TabIndex        =   3
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Creci 37966   -   Rua Bueno Brandão, nº 1500  -  Taboão  -  Guarulhos  -  SP  -  6402-3070"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   7875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "KIELEK IMÓVEIS - Aluga - Vende - Administra"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   4410
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   0
         X2              =   8280
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Contratos:"
      Height          =   1935
      Left            =   240
      TabIndex        =   31
      Top             =   720
      Width           =   8295
      Begin MSDBGrid.DBGrid Grid 
         Bindings        =   "frmRecibos2.frx":1772
         Height          =   1215
         Left            =   120
         OleObjectBlob   =   "frmRecibos2.frx":178C
         TabIndex        =   38
         Top             =   600
         Width           =   8055
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Text            =   "Digite o nome !!!"
         Top             =   240
         Width           =   8055
      End
   End
End
Attribute VB_Name = "frmRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Db As DAO.Database
Dim Tabl As DAO.Recordset
Dim CN As New ADODB.Connection
Dim Sql As String

Private Sub Command1_Click()

frmVencimentos.Text7 = Label31
frmVencimentos.Text1 = Label15
frmVencimentos.Text2 = Label16
frmVencimentos.Text3 = Label17
frmVencimentos.Text4 = Label26
frmVencimentos.Text5 = Label25
frmVencimentos.Text6 = dtaContrato.Recordset("Bimovel")
frmVencimentos.Label12 = CDate(Label23)
frmVencimentos.Label14 = CDate(Label29)
Verifica
frmVencimentos.Show 1


End Sub

Private Sub Command2_Click()
Dim Linha As Integer
    Linha = 1
    
    If Linha = 1 Then
        Cabecalho
    End If
    If Linha = 50 Then
        Printer.NewPage
        Linha = 1
    End If
    
    Printer.EndDoc
    
Command2.Enabled = False

MsgBox ("Os dados foram enviados para impressora!")

If MsgBox("Quer imprimir mais uma VIA?", vbYesNo, "Imprimir") = vbYes Then
    Linha = 1
    
    If Linha = 1 Then
        Cabecalho
    End If
    If Linha = 50 Then
        Printer.NewPage
        Linha = 1
    End If
    Printer.EndDoc
    MsgBox ("Os dados foram enviados para impressora!")
Else
    FechaLabels
    Text2.Visible = False
    Text1 = Empty
    Command2.Enabled = False
    Command3.Caption = "S&air"
    Command1.Enabled = False
    Command1.Caption = "Ger&ar Recibos"
    Command4.Visible = False
    Command5.Enabled = False
    Command6.Enabled = True
    Command7.Enabled = True
    Command9.Enabled = False
    Label3.Visible = False
    Unload frmVencimentos
    Text2.Visible = False
    Text1 = Empty
    Command2.Enabled = False
    Command3.Caption = "S&air"
    Command1.Enabled = False
    Command1.Caption = "Ger&ar Recibos"
    Command4.Visible = False
    Command5.Enabled = False
    Command6.Enabled = True
    Command7.Enabled = True
    Command9.Enabled = False
    Label3.Visible = False
    Unload frmVencimentos
    Exit Sub
End If
    FechaLabels
    Text2.Visible = False
    Text1 = Empty
    Command2.Enabled = False
    Command3.Caption = "S&air"
    Command1.Enabled = False
    Command1.Caption = "Ger&ar Recibos"
    Command4.Visible = False
    Command5.Enabled = False
    Command6.Enabled = True
    Command7.Enabled = True
    Command9.Enabled = False
    Label3.Visible = False
    Unload frmVencimentos
End Sub

Private Sub Command3_Click()

If Command3.Caption = "C&ancelar" Then
    FechaLabels
    Text2.Visible = False
    Text1 = Empty
    Command2.Enabled = False
    Command3.Caption = "S&air"
    Command1.Enabled = False
    Command1.Caption = "Ger&ar Recibos"
    Command4.Visible = False
    Command5.Enabled = False
    Command6.Enabled = True
    Command7.Enabled = True
    Command9.Enabled = False
    Label3.Visible = False
    Unload frmVencimentos
Else
    If MsgBox("Quer sair do Gera Recibos?", vbYesNo, "Sair do Gerador") = vbYes Then
    Unload Me
    RedefineFormPrincipal
Else
    Exit Sub
End If
End If
End Sub

Private Sub Command4_Click()
Text2 = ""
Text2.SetFocus
End Sub

Private Sub Command5_Click()

frmAlteraDados.Text1 = Label15
frmAlteraDados.Text2 = Label16
frmAlteraDados.Text3 = Label17
frmAlteraDados.Text4 = Label18 & "  À  " & Label32
frmAlteraDados.Text5 = Label25
frmAlteraDados.Text6 = Label26
frmAlteraDados.Text7 = Label27
frmAlteraDados.Show 1

End Sub

Private Sub Command6_Click()
If Command3.Caption = "C&ancelar" Then
    FechaLabels
    Text2.Visible = False
    Text1 = Empty
    Command2.Enabled = False
    Command3.Caption = "S&air"
    Command1.Enabled = False
    Command1.Caption = "Ger&ar Recibos"
    Command4.Visible = False
    Command5.Enabled = False
    Command6.Enabled = True
    Command7.Enabled = True
    Command9.Enabled = False
    Label3.Visible = False
    Unload frmVencimentos
End If
Command2.Enabled = True
Command3.Caption = "C&ancelar"
Command6.Enabled = False
FechaLabels
Text2.Visible = True
Text2.SetFocus
Command2.Enabled = True
End Sub

Private Sub Command7_Click()

Label25 = "0,00"
Label26 = "0,00"
Label27 = "0,00"
Label28 = "0,00"

Command1.Enabled = True
Command2.Enabled = False
Command3.Caption = "C&ancelar"
AbreLabels
LabelLoc
Frame5.Enabled = True
Command7.Enabled = False
Text1.Enabled = True
Text1.SetFocus

End Sub

Private Sub Command8_Click()

Command8.Enabled = False
Command2.Enabled = False
Command3.Caption = "C&ancelar"
dtaContrato.DatabaseName = ""
dtaContrato.DatabaseName = App.Path & "\dados\bdimobiliaria.MDB"
dtaContrato.RecordSource = "AContrato"
Frame5.Enabled = True
LabelVenda
AbreLabels
ConfigLabel
dtaContrato.Refresh

End Sub

Private Sub Command9_Click()

Label28 = Format(CDbl(Label25) + (Label26) + (Label27), "###,##0.00")

End Sub

Private Sub Form_Load()

dtaContrato.DatabaseName = App.Path & "\Dados\Bdimobiliaria.MDB"
dtaContrato.RecordSource = "Contrato"

Label15.Visible = False
Label16.Visible = False
Label17.Visible = False

FechaLabels
Label1.Visible = True
Label2.Visible = True
Label3.Visible = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmFundo.Enabled = True
RedefineFormPrincipal
End Sub

Private Sub Label23_Change()
On Error Resume Next
Label23 = CDate(Label23)
End Sub

Private Sub Label29_Change()
On Error Resume Next
Label29 = CDate(Label29)
End Sub

Private Sub Text1_GotFocus()
Text1 = ""
Text1.BackColor = &HFFFF&
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Text1.Text = "" Then
    dtaContrato.RecordSource = "SELECT * FROM CONTRATO"
    dtaContrato.Refresh
    Exit Sub
End If

dtaContrato.RecordSource = "SELECT * FROM CONTRATO WHERE Locador Like '" & Text1.Text & "*'"
dtaContrato.Refresh
End Sub

Private Sub Text1_LostFocus()
Text1.BackColor = &H80000005

If Text1.Text = "" Then
    Text1.Text = "Digite o nome !!!"
End If
End Sub

Private Sub Text3_GotFocus()
Text3.BackColor = &HFFFF&
End Sub

Private Sub Text3_LostFocus()
Text3.BackColor = &H80000005
End Sub

Private Sub Text2_Change()
If Text2 = "" Then
    Command4.Visible = False
    Command4.Enabled = False
Else
    Command4.Visible = True
    Command4.Enabled = True
End If
End Sub

Private Sub Text2_GotFocus()
Text2.BackColor = &HFFFF&
End Sub

Private Sub Text2_LostFocus()
Text2.BackColor = &H80000005
End Sub

Private Function FechaLabels()
Dim LabelObjeto As Object

For Each LabelObjeto In Me.Controls
If TypeOf LabelObjeto Is Label Then
    LabelObjeto.Visible = False
End If
Next LabelObjeto

Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
End Function

Private Function AbreLabels()
Dim LabelObjeto As Object

For Each LabelObjeto In Me.Controls
If TypeOf LabelObjeto Is Label Then
    LabelObjeto.Visible = True
End If
Next LabelObjeto

Command4.Visible = False
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label32.Visible = False

End Function

Private Function LabelLoc()

Label3.Caption = "RECIBO DE LOCAÇÃO"
Label4.Caption = "LOCADOR............:"
Label5.Caption = "LOCATÁRIO........:"
Label6.Caption = "END./BAIRRO......:"
Label7.Caption = "DESCRIÇÃO"
Label8.Caption = "ALUGUÉL DE "
Label9.Caption = "I.P.T.U.............................................................:"
Label10.Caption = "MULTA (10%)................................................:"
Label11.Caption = "TOTAL............................................................:"
Label12.Caption = "OBS.:"
Label13.Caption = "INÍCIO CONTRATO.......:"
Label14.Caption = "TÉRMINO CONTRATO..:"
Label22.Caption = "CÓDIGO:"
Label24.Caption = "ALUGUÉL: 00/00"

End Function

Private Function LabelVenda()

Label3.Caption = "RECIBO DE PRESTAÇÃO"
Label4.Caption = "VENDEDOR..............:"
Label5.Caption = "COMPRADOR...........:"
Label6.Caption = "END./BAIRRO.......:"
Label7.Caption = "DESCRIÇÃO"
Label8.Caption = "PRESTAÇÃO DE........:"
Label9.Caption = "..........................:"
Label10.Caption = "MULTA (10%).........:"
Label11.Caption = "TOTAL...................:"
Label12.Caption = "OBS.:"
Label13.Caption = "INÍCIO PRESTAÇÃO...:"
Label14.Caption = "TÉRMINO PRESTAÇÃO..:"
Label22.Caption = "CÓDIGO:"
Label24.Caption = "PARCELA: 00/00"

End Function

Private Function Verifica()
On Error Resume Next
Dim Cod As String

frmVencimentos.Data2.RecordSource = "SELECT * FROM VENCIMENTOS WHERE LOCATARIO Like '" & Label16 & "*'"
frmVencimentos.Data2.Refresh

Cod = frmVencimentos.Data2.Recordset("Locatario")
If Cod = Label16.Caption Then
    frmVencimentos.Label16.Caption = "1"
    frmVencimentos.Text8 = Label30
    frmVencimentos.Text8.Locked = True
Else
    frmVencimentos.Label16.Caption = "0"
    frmVencimentos.Text8 = ""
    frmVencimentos.Text8.Locked = False
End If

End Function

Private Function Cabecalho()

    Printer.Font = "TIMES NEW ROMAN"
    Printer.FontBold = True
    Printer.FontSize = 14
    
    Printer.Print Tab(5); "-----------------------------------------------------------------------------------------------------------------"
    Printer.Print Tab(5); Label1
    Printer.Print Tab(5); Label2
    Printer.Print Tab(5); "-----------------------------------------------------------------------------------------------------------------"
    
    Printer.Print
    
    Printer.Print Tab(40); Label3
    
    Printer.Font = "TIMES NEW ROMAN"
    Printer.FontSize = 12
    Printer.Print Tab(5); Label22 & "    " & "KC" & Format(Label31, "000")
    Printer.Print Tab(5); Label4 & "        " & Label15
    Printer.Print Tab(5); Label5 & "        " & Label16
    Printer.Print Tab(5); Label6 & "        " & Label17
    Printer.Print
    Printer.Print
    Printer.Print Tab(5); Label8 & "        " & Label18 & "  À  " & Label32 & "     " & "   " & "R$ " & Label25
    Printer.Print Tab(5); Label9 & "         " & "R$ " & Label26
    Printer.Print Tab(5); Label10 & "        " & "R$ " & Label27
    Printer.Print Tab(5); Label11 & "        " & "R$ " & Label28
    Printer.Print Tab(5); Label12
    Printer.Print
    Printer.Print Tab(5); Label13 & "        " & Label23
    Printer.Print Tab(5); Label14 & "        " & Label29 & "       " & Label24
    Printer.Font = "TIMES NEW ROMAN"
    Printer.FontBold = True
    Printer.FontSize = 14
    Printer.Print Tab(5); "-----------------------------------------------------------------------------------------------------------------"
    
    Printer.Font = "TIMES NEW ROMAN"
    Printer.FontSize = 11
    Printer.Print Tab(5); "A QUITAÇÃO DE RECIBO PAGO COM CHEQUE SO SE EFETIVARÁ APÓS A SUA LIQUIDAÇÃO, NÃO É"
    Printer.Print Tab(5); "VÁLIDO RECIBO SEM CARIMBO E ASSINATURA, E O MESMO NÃO QUITA DÉBITOS ANTERIORES."
    Printer.Print
    Printer.Print Tab(5); "--                                                                                                             --"
    Printer.FontBold = False
    
    Printer.Print
    Printer.Font = "TIMES NEW ROMAN"
    Printer.FontBold = True
    Printer.FontSize = 14
    
    Printer.Print Tab(5); "-----------------------------------------------------------------------------------------------------------------"
    Printer.Print Tab(5); Label1
    Printer.Print Tab(5); Label2
    Printer.Print Tab(5); "-----------------------------------------------------------------------------------------------------------------"
    
    Printer.Print
    
    Printer.Print Tab(40); Label3
    
    Printer.Font = "TIMES NEW ROMAN"
    Printer.FontSize = 12
    Printer.Print Tab(5); Label22 & "    " & "KC" & Format(Label31, "000")
    Printer.Print Tab(5); Label4 & "        " & Label15
    Printer.Print Tab(5); Label5 & "        " & Label16
    Printer.Print Tab(5); Label6 & "        " & Label17
    Printer.Print
    Printer.Print
    Printer.Print Tab(5); Label8 & "        " & Label18 & "  À  " & Label32 & "     " & "   " & "R$ " & Label25
    Printer.Print Tab(5); Label9 & "         " & "R$ " & Label26
    Printer.Print Tab(5); Label10 & "        " & "R$ " & Label27
    Printer.Print Tab(5); Label11 & "        " & "R$ " & Label28
    Printer.Print Tab(5); Label12
    Printer.Print
    Printer.Print Tab(5); Label13 & "        " & Label23
    Printer.Print Tab(5); Label14 & "        " & Label29 & "       " & Label24
    Printer.Font = "TIMES NEW ROMAN"
    Printer.FontBold = True
    Printer.FontSize = 14
    Printer.Print Tab(5); "-----------------------------------------------------------------------------------------------------------------"
    
    Printer.Font = "TIMES NEW ROMAN"
    Printer.FontSize = 11
    Printer.Print Tab(5); "DECLARO TER RECEBIDO DA IMOBILIÁRIA KIELEK IMÓVEIS A QUANTIA DE R$_____________"

End Function

Private Function ConfigLabel()

Label31.DataField = codigo
Label31.DataField = vendedor
Label31.DataField = comprador

End Function

Private Function FechaCaixas()
Dim P_objeto As Object

For Each P_objeto In frmVencimentos.Controls
    If TypeOf P_objeto Is TextBox Then
    P_objeto.Locked = True
    P_objeto.BackColor = &H80000013
End If
Next P_objeto

End Function

Private Function RedefineFormPrincipal()

If frmFundo.Text1.Text = "Basico" Then
    frmFundo.cmdCadastros.Enabled = True
    frmFundo.cmdContratos.Enabled = True
    frmFundo.cmdRecibos.Enabled = False
    frmFundo.cmdPrest.Enabled = False
    frmFundo.cmdSair.Caption = "Sair do Programa"
End If
If frmFundo.Text1.Text = "Intermediario" Then
    frmFundo.cmdCadastros.Enabled = True
    frmFundo.cmdContratos.Enabled = True
    frmFundo.cmdRecibos.Enabled = True
    frmFundo.cmdPrest.Enabled = True
    frmFundo.cmdAlugueis.Enabled = True
    frmFundo.cmdSair.Caption = "Sair do Programa"
End If
If frmFundo.Text1.Text = "Avançado" Then
    frmFundo.cmdCadastros.Enabled = True
    frmFundo.cmdContratos.Enabled = True
    frmFundo.cmdRecibos.Enabled = True
    frmFundo.cmdPrest.Enabled = True
    frmFundo.cmdAlugueis.Enabled = True
    frmFundo.cmdSair.Caption = "Sair do Programa"
End If
End Function
