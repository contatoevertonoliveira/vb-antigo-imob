VERSION 5.00
Begin VB.Form frmRecibos 
   Caption         =   "Super Imob - Emissão de Recibos"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   Icon            =   "frmRecibos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   8295
      Begin VB.CommandButton Command3 
         Caption         =   "Sa&ir"
         Height          =   855
         Left            =   6600
         Picture         =   "frmRecibos.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Imprimir Recibo"
         Height          =   855
         Left            =   6600
         Picture         =   "frmRecibos.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ger&ar Recibo"
         Height          =   855
         Left            =   6600
         Picture         =   "frmRecibos.frx":0EEE
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Label29"
         Height          =   195
         Left            =   3240
         TabIndex        =   36
         Top             =   3480
         Width           =   570
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Label28"
         Height          =   195
         Left            =   3240
         TabIndex        =   35
         Top             =   3240
         Width           =   570
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Label27"
         Height          =   195
         Left            =   3240
         TabIndex        =   34
         Top             =   3000
         Width           =   570
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Label26"
         Height          =   195
         Left            =   3240
         TabIndex        =   33
         Top             =   2760
         Width           =   570
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Label25"
         Height          =   195
         Left            =   3240
         TabIndex        =   32
         Top             =   2520
         Width           =   570
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Label24"
         Height          =   195
         Left            =   1680
         TabIndex        =   28
         Top             =   4080
         Width           =   570
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Label23"
         Height          =   195
         Left            =   1680
         TabIndex        =   27
         Top             =   3840
         Width           =   570
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Label22"
         Height          =   195
         Left            =   1680
         TabIndex        =   26
         Top             =   3480
         Width           =   570
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Label21"
         Height          =   195
         Left            =   1680
         TabIndex        =   25
         Top             =   3240
         Width           =   570
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Label20"
         Height          =   195
         Left            =   1680
         TabIndex        =   24
         Top             =   3000
         Width           =   570
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Label19"
         Height          =   195
         Left            =   1680
         TabIndex        =   23
         Top             =   2760
         Width           =   570
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Label18"
         Height          =   195
         Left            =   1680
         TabIndex        =   22
         Top             =   2520
         Width           =   570
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Label17"
         Height          =   195
         Left            =   1680
         TabIndex        =   21
         Top             =   1680
         Width           =   570
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Label16"
         Height          =   195
         Left            =   1680
         TabIndex        =   20
         Top             =   1440
         Width           =   570
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Label15"
         Height          =   195
         Left            =   1680
         TabIndex        =   19
         Top             =   1200
         Width           =   570
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Label14"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   4080
         Width           =   570
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Label13"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   3840
         Width           =   570
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Label12"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   3480
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Label11"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   3240
         Width           =   570
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Label10"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   3000
         Width           =   570
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Label9"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   2760
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Label8"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Label7"
         Height          =   195
         Left            =   1680
         TabIndex        =   11
         Top             =   2160
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Label6"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Label5"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Label4"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   195
         Left            =   3480
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
   Begin VB.Frame Frame1 
      Caption         =   "Selecione a opção de recibo:"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8295
      Begin VB.OptionButton Option3 
         Caption         =   "Recibo Universal"
         Height          =   195
         Left            =   6480
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Recibo de Compra e Venda"
         Height          =   195
         Left            =   3000
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Recibo de Locação"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
If MsgBox("Quer sair do Gera Recibos?", vbYesNo, "Sair do Gerador") = vbYes Then
    Unload Me
Else
    Exit Sub
End If
End Sub

Private Sub Option1_Click()
If Option1.Value Then
    Label3.Caption = "RECIBO DE ALUGUÉL"
    Label4.Caption = "LOCADOR......:"
    Label5.Caption = "LOCATÁRIO....:"
    Label6.Caption = "ENDEREÇO.....:"
    Label8.Caption = "VENCIMENTO DE:"
    Label9.Caption = "I.P.T.U......:"
    Label10.Caption = "MULTA........:"
    Label11.Caption = "TOTAL........:"
    Label12.Caption = "OBS..........:"
End If
End Sub

Private Sub Option2_Click()
    Label3.Caption = "RECIBO DE PRESTAÇÃO"
    Label4.Caption = "VENDEDOR.....:"
    Label5.Caption = "COMPRADOR....:"
    Label6.Caption = "ENDEREÇO.....:"
    Label8.Caption = "VENCIMENTO DE:"
    Label9.Caption = "I.P.T.U......:"
    Label10.Caption = "MULTA........:"
    Label11.Caption = "TOTAL........:"
    Label12.Caption = "OBS..........:"
End Sub
