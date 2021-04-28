Attribute VB_Name = "Module2"
Option Explicit

Public CONE As ADODB.Connection
Public RSTB As ADODB.Recordset

Sub Connection()

Set CONE = CreateObject("ADODB.Connection")
Set RSTB = CreateObject("ADODB.Recordset")

CONE.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source =" & App.Path & "\Dados\bdimobiliaria.mdb;"
RSTB.CursorLocation = adUseClient

End Sub
Sub Disconnection()

RSTB.Close
CONE.Close

Set RSTB = Nothing
Set CONE = Nothing

End Sub
