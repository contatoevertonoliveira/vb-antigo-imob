VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRetconsulta2 
   Caption         =   "Consulta de Clientes Cadastrados"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmRetconsulta2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sa&ir"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   7560
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid1 
      Height          =   6615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   11668
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRetconsulta2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bd As Database
Dim Tb As Recordset

Private Sub cmdSair_Click()
If MsgBox("Quer sair da Busca?", vbYesNo, "Sair da Busca") = vbYes Then
    Unload Me
    Bd.Close
Else
    Exit Sub
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

Set Bd = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\Dados\bdimobiliaria.mdb")
Set Tb = Bd.OpenRecordset("Loc", dbOpenTable)
Tb.Index = "Indcod"

MGrid1.AllowUserResizing = flexResizeBoth


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
     
          
          MGrid1.Rows = MGrid1.Rows + 1

          
          MGrid1.Row = MGrid1.Rows - 1
                  
                    
                    
          For i = 0 To Tb.Fields.Count - 1
               
               
               MGrid1.Col = i + 1
               MGrid1.text = Tb(i).Value & ""
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

End Sub

Private Sub MGrid1_Click()
If MGrid1.Col <> 1 Then
MGrid1.Col = 1
End If

frmClientes.txtcod.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtlocatario.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtnacional.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.TxtProfissao.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.CobEstadoCivil.AddItem MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtregime.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtruanot.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtbairronotifi.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtcomplenot.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtnnot.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtnumAP.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtbloco.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtcidadenotifi.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtest.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtcepnotifi.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txttelefores.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txttelefonecom.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtCpf.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtRg.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtemail.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtsite.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtconjuge.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtconjugeNacional.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtProfissaoConjuge.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtCpfconjuge.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1
frmClientes.txtRgconjuge.text = MGrid1.text
MGrid1.Col = MGrid1.Col + 1

Unload Me

End Sub

