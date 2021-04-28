VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConRec 
   Caption         =   "Super Imob - Gerar Recibo"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   Icon            =   "frmContrato.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   1215
      Left            =   4680
      TabIndex        =   13
      Top             =   5040
      Width           =   4215
      Begin VB.CommandButton cmdVizualiza 
         Caption         =   "Visualiza Recibo"
         Height          =   855
         Left            =   120
         Picture         =   "frmContrato.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "C&ancelar"
         Height          =   855
         Left            =   3120
         Picture         =   "frmContrato.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Vis&ualiza Contrato"
         Height          =   855
         Left            =   1560
         Picture         =   "frmContrato.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resultado:"
      Height          =   1215
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   4335
      Begin VB.TextBox Text2 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Locatário:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proprietário:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Digite o nome do Locatário:"
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   8655
      Begin VB.CommandButton cmdBusca1 
         Caption         =   "B&uscar"
         Height          =   375
         Left            =   6360
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtBusca2 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   5655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Digite o nome do Proprietário:"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   8655
      Begin VB.CommandButton cmdBusca 
         Caption         =   "B&uscar"
         Height          =   375
         Left            =   6360
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtBusca1 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   5655
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid1 
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   2355
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid2 
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      Top             =   3360
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   2566
      _Version        =   393216
   End
End
Attribute VB_Name = "frmConRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bd As Database
Dim Tb As Recordset
Dim Tb2 As Recordset

Private Sub cmdsair_Click()
Unload Me
End Sub

Private Sub cmdVizualiza_Click()
If MGrid1.Col <> 2 Then
MGrid1.Col = 2
End If
Unload frmConRec
frmRecibos.Show
frmRecibos.Label15.Caption = Text1.Text
frmRecibos.Label16.Caption = Text2.Text

End Sub

Private Sub Form_Load()

Set Bd = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\Dados\bdimobiliaria.mdb")
Set Tb = Bd.OpenRecordset("Prop", dbOpenTable)
Set Tb2 = Bd.OpenRecordset("Loc", dbOpenTable)
Tb.Index = "Indcod"
Tb2.Index = "Indcod"

MGrid1.AllowUserResizing = flexResizeBoth


DoEvents
MGrid1.AllowUserResizing = flexResizeBoth



     MGrid1.Cols = Tb.Fields.Count + 1
     
     
     MGrid1.Rows = 1

DoEvents



For X = 0 To MGrid1.Cols - 1
           MGrid1.TextMatrix(0, X) = "Nome"
           DoEvents
Next
        

MGrid1.Row = 0

     For i = 0 To Tb.Fields.Count - 1
          
          MGrid1.Col = i + 1
            
         
          MGrid1.ColAlignment(i + 1) = flexAlignLeftCenter

          MGrid1.ColWidth(i + 1) = 1500
          MGrid1.Text = Tb.Fields(i).Name
         
          DoEvents
          
     Next
MGrid1.ColWidth(4) = 3500

CONTADOR = 0

     Do While Not Tb.EOF
     
          
          MGrid1.Rows = MGrid1.Rows + 1

        
          MGrid1.Row = MGrid1.Rows - 1
          
          For i = 0 To Tb.Fields.Count - 1
               
               MGrid1.Col = i + 1
               MGrid1.Text = Tb(i).Value & ""
             DoEvents
            
          Next
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


 
    With MGrid1
                       
        
         For X = 1 To MGrid1.Rows - 1
           MGrid1.TextMatrix(X, 0) = Str(X)
        Next
    End With


MGrid2.AllowUserResizing = flexResizeBoth


DoEvents
MGrid2.AllowUserResizing = flexResizeBoth



     MGrid2.Cols = Tb2.Fields.Count + 1
     
     
     MGrid2.Rows = 1

DoEvents



For X = 0 To MGrid2.Cols - 1
           MGrid2.TextMatrix(0, X) = "Nome"
           DoEvents
Next
        

MGrid1.Row = 0

     For i = 0 To Tb.Fields.Count - 1
          
          MGrid2.Col = i + 1
            
         
          MGrid2.ColAlignment(i + 1) = flexAlignLeftCenter

          MGrid2.ColWidth(i + 1) = 1500
          MGrid2.Text = Tb2.Fields(i).Name
         
          DoEvents
          
     Next
MGrid2.ColWidth(4) = 3500

CONTADOR = 0

     Do While Not Tb2.EOF
     
          
          MGrid2.Rows = MGrid2.Rows + 1

        
          MGrid2.Row = MGrid2.Rows - 1
          
          For i = 0 To Tb2.Fields.Count - 1
               
               MGrid2.Col = i + 1
               MGrid2.Text = Tb2(i).Value & ""
             DoEvents
            
          Next
          Tb2.MoveNext
     DoEvents

      CONTADOR = CONTADOR + 1
                
        
        
            If CONTADOR = 2000 Then
            If MsgBox("Chegou em 2000 registros" & vbCrLf & "deseja cancelar?", vbYesNo + vbQuestion, "Atenção") = vbYes Then
            GoTo pular2000
            End If
            End If
        
        

  Loop
     
     
pular2000:


 
    With MGrid2
                       
        
         For X = 1 To MGrid2.Rows - 1
           MGrid2.TextMatrix(X, 0) = Str(X)
        Next
    End With


End Sub

Private Sub MGrid1_Click()
If MGrid1.Col <> 2 Then
MGrid1.Col = 2
End If

Text1.Text = MGrid1.Text
MGrid1.Col = MGrid1.Col + 1

End Sub

Private Sub MGrid2_Click()
If MGrid2.Col <> 2 Then
MGrid2.Col = 2
End If

Text2.Text = MGrid2.Text
MGrid2.Col = MGrid2.Col + 1
End Sub
