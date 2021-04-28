VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmVizCon 
   Caption         =   "Super Imob - Visualizar Contratos e Gerar Recibos"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmVizCon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Visualização do contrato selecionado:"
      Height          =   2415
      Left            =   240
      TabIndex        =   9
      Top             =   3240
      Width           =   11415
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   360
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command4 
      Height          =   615
      Left            =   10440
      Picture         =   "frmVizCon.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   9240
      Picture         =   "frmVizCon.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   8040
      Picture         =   "frmVizCon.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   6840
      Picture         =   "frmVizCon.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdBusca 
      Height          =   615
      Left            =   5640
      Picture         =   "frmVizCon.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   450
      Width           =   5055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contratos Gerados:"
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   11415
      Begin MSFlexGridLib.MSFlexGrid MGrid1 
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   2778
         _Version        =   393216
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Digite um nome para a consulta:"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   2280
   End
End
Attribute VB_Name = "frmVizCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
MGrid1.AllowUserResizing = flexResizeBoth
'frmRetCons.Caption = "Consulta tarifas por periodo"
'frmTeste.Caption = "Consulta tarifas por periodo"

DoEvents
MGrid1.AllowUserResizing = flexResizeBoth



     MGrid1.Cols = Tb.Fields.Count + 1
     
     
     MGrid1.Rows = 1

DoEvents



For X = 0 To MGrid1.Cols - 1
           MGrid1.TextMatrix(0, X) = "codigo"
           DoEvents
Next
        

MGrid1.Row = 0

     For i = 0 To Tb.Fields.Count - 1
          
          MGrid1.Col = i + 1
            
         
          MGrid1.ColAlignment(i + 1) = flexAlignLeftCenter

          MGrid1.ColWidth(i + 1) = 1500
          MGrid1.text = Tb.Fields(i).Name
         
          DoEvents
          
     Next
MGrid1.ColWidth(4) = 3500

CONTADOR = 0

     Do While Not Tb.EOF
     
          'Add a row to the FlexGrid everytime
          'the database goes to another row.
          MGrid1.Rows = MGrid1.Rows + 1

          'Move to last row to add data.
          MGrid1.Row = MGrid1.Rows - 1
          
          
                    
          'Move to every cell in the row
          'and fill it in with the
          'corresponding value from the
          'database.
          
          For i = 0 To Tb.Fields.Count - 1
               'Remember that the
               'first column is left blank so
               'we shift over 1.
               
               MGrid1.Col = i + 1
               MGrid1.text = Tb(i).Value & ""
             DoEvents
            
          Next
          'Move to the next record.
          Tb.MoveNext
     DoEvents

      CONTADOR = CONTADOR + 1
                
        
        
            If CONTADOR = 2000 Then
            If MsgBox("Chegou em 2000 registros" & vbCrLf & "deseja cancelar?", vbYesNo + vbQuestion, "Atenção") = vbYes Then
            GoTo pula2000
            End If
            End If
        
        

  Loop
     
     
pula2000:


'PARA NUMERAR LINHAS DA CONSULTA NO FLEXGRID
 
    With MGrid1
'        .ColAlignment(-1) = 1       'all Left alligned
                       
        
         For X = 1 To MGrid1.Rows - 1
           MGrid1.TextMatrix(X, 0) = Str(X)
        Next
'        .Row = 1
'        .Col = 1
'        .CellBackColor = &HC0FFFF   'lt. yellow
    End With
End Sub

Private Sub MGrid1_Click()
If MGrid1.Col <> 1 Then
MGrid1.Col = 1
End If

FrmProprietarios.txtcod.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.txtproprietario.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.txtnacional.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.TxtProfissao.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.CobEstadoCivil.AddItem MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.txtregime.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.txtruanot.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.txtbairronotifi.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.txtcomplenot.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.txtnnot.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.txtnumAP.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.txtbloco.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.txtcidadenotifi.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.txtest.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.txtcepnotifi.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.txttelefores.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.txttelefonecom.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.mskCpf.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.mskRg.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.txtemail.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.txtsite.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.txtconjuge.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.txtconjugeNacional.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.txtProfissaoConjuge.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.mskCpfconjuge.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
FrmProprietarios.mskRgconjuge.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1

Unload Me
End Sub
