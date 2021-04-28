VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBusca2 
   BorderStyle     =   0  'None
   Caption         =   "Buscar Pessoas"
   ClientHeight    =   4590
   ClientLeft      =   510
   ClientTop       =   420
   ClientWidth     =   6420
   Icon            =   "frmBusca2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdsair 
      Caption         =   "&Sair da Pesquisa"
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   6135
      Begin VB.OptionButton opt2 
         Caption         =   "2 - Banco de Dados"
         Height          =   195
         Left            =   3600
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton opt1 
         Caption         =   "1 - Banco de Dados"
         Height          =   195
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Resultado:"
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   6135
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1815
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   3201
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Digite algum nome para buscar:"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton cmdbusca 
         Caption         =   "&Pesquisar"
         Height          =   375
         Left            =   4200
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000017&
      X1              =   360
      X2              =   3960
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000017&
      X1              =   120
      X2              =   3360
      Y1              =   4200
      Y2              =   4200
   End
End
Attribute VB_Name = "frmBusca2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbusca_Click()

If opt1.Value = False And opt2.Value = False Then
    busca = MsgBox("Selecione uma opção abaixo", vbInformation, "Busca de Pessoas")
End If
If opt1.Value = True Then
If cmdbusca.Caption = "&Retornar" Then
        Rs4.Close
        Rs4.Open "select * from clientes ", cnn2
        cmdbusca.Caption = "&Pesquisar"
    Else
        resposta = Text1.Text
        Rs4.Close
        Rs4.Open "select * from clientes where nome like '%" & resposta & "%'", cnn2
        If Rs4.EOF Then
            MsgBox "Não encontrei o nome: " & resposta
            Rs4.Close
            Rs4.Open "select * from clientes ", cnn2
        Else
        cmdbusca.Caption = "&Pesquisar"
            Text1.Text = ""
        End If
End If
PreencherDataGrid2
End If

If opt2.Value = True Then
If cmdbusca.Caption = "&Retornar" Then
        tb.Close
        tb.Open "select * from clientes ", BD
        cmdBuscar.Caption = "&Pesquisar"
    Else
        resposta = Text1.Text
        tb.Close
        tb.Open "select * from clientes where nome like '%" & resposta & "%'", BD
        If tb.EOF Then
            MsgBox "Não encontrei o nome: " & resposta
            tb.Close
            tb.Open "select * from clientes ", BD
        Else
        cmdbusca.Caption = "&Pesquisar"
            Text1.Text = ""
        End If
End If
PreencherDataGrid
End If
End Sub

Private Sub cmdsair_Click()
BD.Close
cnn2.Close
Unload frmBusca2
End Sub

Private Sub Form_Deactivate()
Me.BackColor = &H80000001
Frame1.BackColor = &H80000001
Frame2.BackColor = &H80000001
Frame3.BackColor = &H80000001
End Sub

Private Sub Form_Load()
conectar
conectar2
End Sub

Sub PreencherDataGrid()

DataGrid1.Columns.Add (0)

Set DataGrid1.DataSource = tb
DataGrid1.Refresh
DataGrid1.Columns(0).Width = 752
DataGrid1.Columns(1).Width = 2000
DataGrid1.Columns(2).Width = 1000
DataGrid1.Columns(3).Width = 3000
DataGrid1.Columns(4).Width = 752
DataGrid1.Columns(5).Width = 752
DataGrid1.Columns(6).Width = 752
DataGrid1.Columns(7).Width = 752
DataGrid1.Columns(8).Width = 752
DataGrid1.Columns(9).Width = 2500
DataGrid1.Columns(10).Width = 2500

End Sub

Sub PreencherDataGrid2()

DataGrid1.Columns.Add (0)

Set DataGrid1.DataSource = Rs4
DataGrid1.Refresh
DataGrid1.Columns(0).Width = 752
DataGrid1.Columns(1).Width = 2000
DataGrid1.Columns(2).Width = 1000
DataGrid1.Columns(3).Width = 3000
DataGrid1.Columns(4).Width = 752
DataGrid1.Columns(5).Width = 752
DataGrid1.Columns(6).Width = 752
DataGrid1.Columns(7).Width = 752
DataGrid1.Columns(8).Width = 752
DataGrid1.Columns(9).Width = 2500
DataGrid1.Columns(10).Width = 2500

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

If opt1.Value = False And opt2.Value = False Then
    busca = MsgBox("Selecione uma opção abaixo", vbInformation, "Busca de Pessoas")
End If
If opt1.Value = True Then
If cmdbusca.Caption = "&Retornar" Then
        Rs4.Close
        Rs4.Open "select * from clientes ", cnn2
        cmdbusca.Caption = "&Pesquisar"
    Else
        resposta = Text1.Text
        Rs4.Close
        Rs4.Open "select * from clientes where nome like '%" & resposta & "%'", cnn2
        If Rs4.EOF Then
            MsgBox "Não encontrei o nome: " & resposta
            Rs4.Close
            Rs4.Open "select * from clientes ", cnn2
        Else
        cmdbusca.Caption = "&Pesquisar"
            Text1.Text = ""
        End If
End If
PreencherDataGrid2
End If

If opt2.Value = True Then
If cmdbusca.Caption = "&Retornar" Then
        tb.Close
        tb.Open "select * from clientes ", BD
        cmdBuscar.Caption = "&Pesquisar"
    Else
        resposta = Text1.Text
        tb.Close
        tb.Open "select * from clientes where nome like '%" & resposta & "%'", BD
        If tb.EOF Then
            MsgBox "Não encontrei o nome: " & resposta
            tb.Close
            tb.Open "select * from clientes ", BD
        Else
        cmdbusca.Caption = "&Pesquisar"
            Text1.Text = ""
        End If
End If
PreencherDataGrid
End If
End If
End Sub
