VERSION 5.00
Begin VB.Form Frmlogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "Frmlog.frx":0000
   ScaleHeight     =   2250
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtsenha 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   650
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   1680
   End
   Begin VB.TextBox txtnome 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   650
      MaxLength       =   30
      TabIndex        =   2
      Top             =   600
      Width           =   3000
   End
   Begin VB.Label Lblrestantes 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   230
      Left            =   1830
      TabIndex        =   6
      Top             =   1980
      Width           =   150
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2400
      MouseIcon       =   "Frmlog.frx":2C6B
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "Cancelar"
      Top             =   1950
      Width           =   1250
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   200
      Left            =   2650
      MouseIcon       =   "Frmlog.frx":2F75
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "Efetuar login"
      Top             =   1680
      Width           =   1000
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   465
   End
End
Attribute VB_Name = "Frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tentativas As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
         Label3_Click
    End If

End Sub

Private Sub Form_Load()
Me.Left = 150
Me.Top = 2250
End Sub

Private Sub Label3_Click()

    If txtnome.Text = "" Then
        MsgBox "Informe o nome de usuário", vbInformation, " Atenção!"
        txtnome.SetFocus
        Exit Sub
        
    ElseIf txtsenha.Text = "" Then
        MsgBox "Informe a senha", vbInformation, " Atenção!"
        txtsenha.SetFocus
        Exit Sub
    
    Else
    
    Dim Cnn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    
    Set Cnn = New ADODB.Connection
    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    
    Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\Programa Imobiliária\Dados\Bdimobiliaria.mdb;Jet OLEDB:Database;"
    Rs.Open "Select * From Controle Where Nome='" & txtnome.Text & "' and Senha='" & txtsenha.Text & "'", Cnn, adOpenStatic
        
    frmFundo.Enabled = True
            
            If Rs!Tipo = "BASICO" Then
                frmFundo.cmdCadastros.Enabled = True
                frmFundo.cmdContratos.Enabled = True
                frmFundo.cmdAlugueis.Enabled = False
                frmFundo.cmdPrest.Enabled = False
                frmFundo.cmdRecibos.Enabled = False
                frmFundo.cmdSair.Enabled = True
                frmFundo.Text1.Text = "Basico"
            End If
        
            If Rs!Tipo = "INTERMEDIARIO" Then
                frmFundo.cmdCadastros.Enabled = True
                frmFundo.cmdContratos.Enabled = True
                frmFundo.cmdAlugueis.Enabled = True
                frmFundo.cmdPrest.Enabled = True
                frmFundo.cmdRecibos.Enabled = True
                frmFundo.cmdSair.Enabled = True
                frmFundo.Text1.Text = "Intermediario"
            End If
        
            If Rs!Tipo = "AVANÇADO" Then
                frmFundo.cmdCadastros.Enabled = True
                frmFundo.cmdContratos.Enabled = True
                frmFundo.cmdAlugueis.Enabled = True
                frmFundo.cmdPrest.Enabled = True
                frmFundo.cmdRecibos.Enabled = True
                frmFundo.cmdSair.Enabled = True
                frmFundo.Text1.Text = "Avançado"
            End If
    
            frmFundo.Label1.Caption = "Login " & Rs!Tipo & " de " & StrConv(txtnome, vbUpperCase)
    
    If Rs.RecordCount > 0 Then
      
            
            
        
        Rs.Close
        Set Rs = Nothing
        Cnn.Close
        Set Cnn = Nothing
        Unload Me
    
    Else
    
        MsgBox "Usuário ou senha incorretos", vbInformation, " Atenção!"
        
        txtnome.Text = ""
        txtsenha.Text = ""
        txtnome.SetFocus
        
        Tentativas = Tentativas + 1
    
        Lblrestantes.Caption = Lblrestantes.Caption - 1
    
    If Tentativas = 3 Then

        MsgBox "Você ultrapassou o número de tentativas de acesso, o sistema será fechado", vbCritical, " Atenção!"
        End
        
    End If
    End If
   End If
    
    
End Sub

Private Sub Label4_Click()

    End

End Sub

Private Sub txtnome_GotFocus()

    txtnome.BackColor = &HC0E0FF

End Sub

Private Sub txtnome_LostFocus()

    txtnome.BackColor = &HFFFFFF

End Sub

Private Sub txtsenha_GotFocus()

    txtsenha.BackColor = &HC0E0FF

End Sub

Private Sub txtsenha_LostFocus()

    txtsenha.BackColor = &HFFFFFF

End Sub
