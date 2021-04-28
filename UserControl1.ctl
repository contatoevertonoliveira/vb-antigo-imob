VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   InvisibleAtRuntime=   -1  'True
   Picture         =   "UserControl1.ctx":0000
   ScaleHeight     =   450
   ScaleWidth      =   450
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function Parse(Code As String)
popup = InStr(Code, "$VBX->Popup")
endpr = InStr(Code, "$VBX->EndProgram")
pause = InStr(Code, "$VBX->Pause")
Code = Replace(Code, "$VBX->", "")
If popup = 1 Then
Code = Replace(Code, "Popup('", "")
Code = Replace(Code, "')", "")
Alert Code
End If
If endpr = 1 Then
End
End If
If pause = 1 Then
Alert "USING AN OLD CODE VERSION" & vbNewLine & "RECODE/REINSTALL"
End If
End Function

Private Sub Alert(Text As String)
    ' Display a new alertbox
    Dim AlertBox As frmAlert
    Set AlertBox = New frmAlert
    AlertBox.DisplayAlert Text, 3000
End Sub

Private Sub Timer1_Timer()
Timer1.Tag = "1"
Timer1.Enabled = False
End Sub

