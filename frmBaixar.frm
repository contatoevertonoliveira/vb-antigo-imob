VERSION 5.00
Begin VB.Form frmBaixar 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alterar baixa de vencimento"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "| C&ancelar |"
      Height          =   495
      Left            =   4440
      TabIndex        =   14
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>>>>> B&aixar Recibo >>>>>>"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
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
         Left            =   2880
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Left            =   600
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   4440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
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
         Left            =   6000
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
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
         Left            =   4440
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Left            =   3000
         TabIndex        =   8
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Left            =   600
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
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
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Outra data"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Data de hoje"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proprietário:"
         Height          =   195
         Left            =   2040
         TabIndex        =   22
         Top             =   1320
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ao Proprietário"
         Height          =   195
         Left            =   3120
         TabIndex        =   21
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data de Pagamento"
         Height          =   195
         Left            =   2880
         TabIndex        =   20
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Iptu:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   4440
         TabIndex        =   18
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor R$:"
         Height          =   195
         Left            =   6120
         TabIndex        =   17
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor R$:"
         Height          =   195
         Left            =   4560
         TabIndex        =   16
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   195
         Left            =   2520
         TabIndex        =   15
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Multa:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmBaixar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Dim valor1 As Currency
Dim valor2 As Currency
Dim valor3 As Currency
Dim valor4 As Currency
Dim valor5 As Currency
Dim resultado As Currency

If Len(Text1) > 0 Then
    valor1 = Text1
Else
    Text1 = "0,00"
End If

If Len(Text2) > 0 Then
    valor2 = Text2
Else
    Text2 = "0,00"
End If

If Len(Text3) > 0 Then
    valor3 = Text3
Else
    Text3 = "0,00"
End If

If Len(Text4) > 0 Then
    valor4 = Text4
Else
    Text4 = "0,00"
End If

If Len(Text6) > 0 Then
    valor5 = Text6
Else
    Text6 = "0,00"
End If

If frmAluguel.Visible = True Then

If Text3.Text = "" Then
    If Option1.Value = True Then
        resultado = valor1 + valor2 + valor5
        Text2 = Format$(resultado, "Currency")
        MsgBox ("O Valor total foi alterado!")
            If MsgBox("Tem certeza que deseja dar baixa nesse lançamento ?" & "Valor: " & Text2.Text, vbQuestion + vbYesNo, "Confirmação") = vbYes Then
                frmAluguel.Data1.Recordset.Edit
                frmAluguel.Data1.Recordset("Recebidos") = Date
                frmAluguel.Data1.Recordset("Multa") = Text1
                frmAluguel.Data1.Recordset("Valor") = Text2
                frmAluguel.Data1.Recordset("Iptu") = Text6
                frmAluguel.Data1.Recordset("Ob1") = Text5
                frmAluguel.Data1.Recordset("Prop") = Text7
                frmAluguel.Data1.Recordset("ValorProp") = Text8
                frmAluguel.Data1.Recordset.Update
                MsgBox ("Seu lançamento foi dado baixa com sucesso !!"), vbInformation, "Sucesso !"
            End If
            Unload Me
        End If
        
        
    If Option2.Value = True Then
        resultado = valor1 + valor2 + valor5
        Text2 = Format$(resultado, "Currency")
        MsgBox ("O Valor total foi alterado!")
            If MsgBox("Tem certeza que deseja dar baixa nesse lançamento ?" & "Valor: " & Text2.Text, vbQuestion + vbYesNo, "Confirmação") = vbYes Then
                frmAluguel.Data1.Recordset.Edit
                frmAluguel.Data1.Recordset("Recebidos") = txtData.Text
                frmAluguel.Data1.Recordset("Multa") = Text1
                frmAluguel.Data1.Recordset("Valor") = Text2
                frmAluguel.Data1.Recordset("Iptu") = Text6
                frmAluguel.Data1.Recordset("Ob1") = Text5
                frmAluguel.Data1.Recordset("Prop") = Text7
                frmAluguel.Data1.Recordset("ValorProp") = Text8
                frmAluguel.Data1.Recordset.Update
                MsgBox ("Seu lançamento foi dado baixa com sucesso !!"), vbInformation, "Sucesso !"
            End If
            Unload Me
        End If
Else
    resultado = valor1 + valor2 + valor3 + valor4 + valor5
    Text2 = Format$(resultado, "Currency")
    MsgBox ("Valor total Atualizado!")
    
    If MsgBox("O Valor foi do Total foi alterado você deseja dar baixa nesse vencimento?" & "Valor: " & Text2.Text, vbYesNo, "Baixar Vencimento") = vbYes Then
        If Option1.Value = True Then
            If MsgBox("Tem certeza que deseja dar baixa nesse lançamento ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
                frmAluguel.Data1.Recordset.Edit
                frmAluguel.Data1.Recordset("Recebidos") = Date
                frmAluguel.Data1.Recordset("Multa") = Text1
                frmAluguel.Data1.Recordset("Valor") = Text2
                frmAluguel.Data1.Recordset("Ob1") = Text5
                frmAluguel.Data1.Recordset("Prop") = Text7
                frmAluguel.Data1.Recordset("ValorProp") = Text8
                frmAluguel.Data1.Recordset.Update
                MsgBox ("Seu lançamento foi dado baixa com sucesso !!"), vbInformation, "Sucesso !"
            End If
            Unload Me
        End If


        If Option2.Value = True Then
            If MsgBox("Tem certeza que deseja dar baixa nesse lançamento ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
                frmAluguel.Data1.Recordset.Edit
                frmAluguel.Data1.Recordset("Recebidos") = txtData.Text
                frmAluguel.Data1.Recordset("Multa") = Text1
                frmAluguel.Data1.Recordset("Valor") = Text2
                frmAluguel.Data1.Recordset("Iptu") = Text6
                frmAluguel.Data1.Recordset("Ob1") = Text5
                frmAluguel.Data1.Recordset("Prop") = Text7
                frmAluguel.Data1.Recordset("ValorProp") = Text8
                frmAluguel.Data1.Recordset.Update
                MsgBox ("Seu lançamento foi dado baixa com sucesso !!"), vbInformation, "Sucesso !"
            End If
            Unload Me
        End If
Else
    Exit Sub
End If
End If

ElseIf frmPrestacao.Visible = True Then

If Text3.Text = "" Then
    If Option1.Value = True Then
        resultado = valor1 + valor2 + valor5
        Text2 = Format$(resultado, "Currency")
        MsgBox ("O Valor total foi alterado!")
            If MsgBox("Tem certeza que deseja dar baixa nesse lançamento ?" & "Valor: " & Text2.Text, vbQuestion + vbYesNo, "Confirmação") = vbYes Then
                frmPrestacao.Data1.Recordset.Edit
                frmPrestacao.Data1.Recordset("Recebidos") = Date
                frmPrestacao.Data1.Recordset("Multa") = Text1
                frmPrestacao.Data1.Recordset("Valor") = Text2
                frmPrestacao.Data1.Recordset("Ob1") = Text5
                frmPrestacao.Data1.Recordset("Prop") = Text7
                frmPrestacao.Data1.Recordset("ValorProp") = Text8
                frmPrestacao.Data1.Recordset.Update
                MsgBox ("Seu lançamento foi dado baixa com sucesso !!"), vbInformation, "Sucesso !"
            End If
            Unload Me
        End If
        
        
    If Option2.Value = True Then
        resultado = valor1 + valor2 + valor5
        Text2 = Format$(resultado, "Currency")
        MsgBox ("O Valor total foi alterado!")
            If MsgBox("Tem certeza que deseja dar baixa nesse lançamento ?" & "Valor: " & Text2.Text, vbQuestion + vbYesNo, "Confirmação") = vbYes Then
                frmPrestacao.Data1.Recordset.Edit
                frmPrestacao.Data1.Recordset("Recebidos") = txtData.Text
                frmPrestacao.Data1.Recordset("Multa") = Text1
                frmPrestacao.Data1.Recordset("Valor") = Text2
                frmPrestacao.Data1.Recordset("Ob1") = Text5
                frmPrestacao.Data1.Recordset("Prop") = Text7
                frmPrestacao.Data1.Recordset("ValorProp") = Text8
                frmPrestacao.Data1.Recordset.Update
                MsgBox ("Seu lançamento foi dado baixa com sucesso !!"), vbInformation, "Sucesso !"
            End If
            Unload Me
        End If
Else
    resultado = valor1 + valor2 + valor3 + valor4 + valor5
    Text2 = Format$(resultado, "Currency")
    MsgBox ("Valor total Atualizado!")
    
    If MsgBox("O Valor foi do Total foi alterado você deseja dar baixa nesse vencimento?" & "Valor: " & Text2.Text, vbYesNo, "Baixar Vencimento") = vbYes Then
        If Option1.Value = True Then
            If MsgBox("Tem certeza que deseja dar baixa nesse lançamento ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
                frmPrestacao.Data1.Recordset.Edit
                frmPrestacao.Data1.Recordset("Recebidos") = Date
                frmPrestacao.Data1.Recordset("Multa") = Text1
                frmPrestacao.Data1.Recordset("Valor") = Text2
                frmPrestacao.Data1.Recordset("Ob1") = Text5
                frmPrestacao.Data1.Recordset("Prop") = Text7
                frmPrestacao.Data1.Recordset("ValorProp") = Text8
                frmPrestacao.Data1.Recordset.Update
                MsgBox ("Seu lançamento foi dado baixa com sucesso !!"), vbInformation, "Sucesso !"
            End If
            Unload Me
        End If


        If Option2.Value = True Then
            If MsgBox("Tem certeza que deseja dar baixa nesse lançamento ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
                frmPrestacao.Data1.Recordset.Edit
                frmPrestacao.Data1.Recordset("Recebidos") = txtData.Text
                frmPrestacao.Data1.Recordset("Multa") = Text1
                frmPrestacao.Data1.Recordset("Valor") = Text2
                frmPrestacao.Data1.Recordset("Ob1") = Text5
                frmPrestacao.Data1.Recordset("Prop") = Text7
                frmPrestacao.Data1.Recordset("ValorProp") = Text8
                frmPrestacao.Data1.Recordset.Update
                MsgBox ("Seu lançamento foi dado baixa com sucesso !!"), vbInformation, "Sucesso !"
            End If
            Unload Me
        End If
Else
    Exit Sub
End If
End If
End If
End Sub

Private Sub Command2_Click()
txtData = ""
Option1.Value = False
Option2.Value = False
Unload Me
End Sub

Private Sub Option1_Click()
txtData = ""
txtData.Enabled = False
Text1 = "0,00"
Text1.Enabled = False
Text3 = "0,00"
Text4 = "0,00"

Text6 = frmAluguel.Data1.Recordset.Fields(5)
Text2 = frmAluguel.Data1.Recordset.Fields(6)
End Sub

Private Sub Option2_Click()
txtData.Enabled = True
Text1.Enabled = True
txtData.SetFocus
End Sub


Private Function Verifica()
Dim valor1 As Currency
Dim valor2 As Currency
Dim valor3 As Currency
Dim valor4 As Currency
Dim valor5 As Currency

If Len(Text1) > 0 Then
    valor1 = Text1
Else
    Text1 = "0,00"
End If

If Len(Text2) > 0 Then
    valor2 = Text2
Else
    Text2 = "0,00"
End If

If Len(Text3) > 0 Then
    valor3 = Text3
Else
    Text3 = "0,00"
End If

If Len(Text4) > 0 Then
    valor4 = Text4
Else
    Text4 = "0,00"
End If

If Len(Text6) > 0 Then
    valor5 = Text6
Else
    Text6 = "0,00"
End If
End Function
