Attribute VB_Name = "MdlWord"
Option Explicit
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
'para esta função abaixo funcionar, vc tem que fazer referências
'para "Microsoft Scripting Runtime"
Public Sub CopyFile(sOrigem As String, sDestino As String)
    On Error Resume Next
    Dim fso As New FileSystemObject
    fso.CopyFile sOrigem, sDestino, True
End Sub
