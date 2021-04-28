Attribute VB_Name = "Module5"
Option Explicit

Public CON As ADODB.Connection
Public RS6 As ADODB.Recordset
Public RS7 As ADODB.Recordset

Sub Connect()

Set CON = CreateObject("ADODB.Connection")
Set RS6 = CreateObject("ADODB.Recordset")
Set RS7 = CreateObject("ADODB.Recordset")

CON.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source =" & App.Path & "\Dados\bdimobiliaria.mdb;"
RS6.CursorLocation = adUseClient
RS7.CursorLocation = adUseClient

End Sub
Sub Disconnect()

RS6.Close
CON.Close

Set RS6 = Nothing
Set CON = Nothing

End Sub

Sub Disconnect2()

RS7.Close
CON.Close

Set RS7 = Nothing
Set CON = Nothing

End Sub
