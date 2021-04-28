Attribute VB_Name = "Module6"
Global Ext
Global Trilhao, Bilhao, Milhao
Global Mil, Unid, Centavos, ValChr

Function Extenso(Valor)

On Error GoTo erro

Dim Trilhao, Bilhao, Milhao, Mil, Unid, Centavos, Ext, ValChr
Dim NmMoe_1sg As String
Dim NmMoe_1pl As String
Dim NmMoe_2sg As String
Dim NmMoe_2pl As String

Dim Comprimento
NmMoe_1sg = "Real"
NmMoe_1pl = "Reais"
NmMoe_2sg = "Centavo"
NmMoe_2pl = "Centavos"
If Len(Valor) = 0 Then
   MsgBox "Nenhum valor para extenso", 16, "Erro nos dados"
   Exit Function
End If

If Len(Valor) > 18 Then
   msg = msg + "Infelizmente , essa função não retorna extenso" + Chr(13)
   msg = msg + "de valores maiores que : 999 trilhões ."
   MsgBox msg, 16, "Estouro de capacidade"
   Exit Function
End If
ValChr = (CCur(Valor))

If Mid(Right(ValChr, 3), 1, 1) = "," Then
   Centavos = Val(Right$(ValChr, 2))
   Comprimento = Len(ValChr)
Else
   If Mid(Right(ValChr, 2), 1, 1) = "," Then
      Centavos = Val(Right$(ValChr, 1) + "0")
      Comprimento = Len(ValChr) + 1
   Else
       Centavos = Val("00")
       Comprimento = Len(ValChr)
   End If
End If

Select Case Comprimento

       Case Is = 18 And Centavos <> 0
           Trilhao = Val(Left$(ValChr, 3))
           Bilhao = Mid$(ValChr, 4, 3)
           Milhao = Val(Mid$(ValChr, 7, 3))
           Mil = Val(Mid$(ValChr, 10, 3))
           Unid = Val(Mid$(ValChr, 13, 3))
       
       Case Is = 17 And Centavos <> 0
           Trilhao = Val(Left$(ValChr, 2))
           Bilhao = Mid$(ValChr, 3, 3)
           Milhao = Val(Mid$(ValChr, 6, 3))
           Mil = Val(Mid$(ValChr, 9, 3))
           Unid = Val(Mid$(ValChr, 12, 3))
 
       Case Is = 16 And Centavos <> 0
           Trilhao = Val(Left$(ValChr, 1))
           Bilhao = Mid$(ValChr, 2, 3)
           Milhao = Val(Mid$(ValChr, 5, 3))
           Mil = Val(Mid$(ValChr, 8, 3))
           Unid = Val(Mid$(ValChr, 11, 3))
       
       Case Is = 15 And Centavos <> 0
           Bilhao = Mid$(ValChr, 1, 3)
           Milhao = Val(Mid$(ValChr, 4, 3))
           Mil = Val(Mid$(ValChr, 7, 3))
           Unid = Val(Mid$(ValChr, 10, 3))
       
       Case Is = 14 And Centavos <> 0
           Bilhao = Mid$(ValChr, 1, 2)
           Milhao = Val(Mid$(ValChr, 3, 3))
           Mil = Val(Mid$(ValChr, 6, 3))
           Unid = Val(Mid$(ValChr, 9, 3))
       
       Case Is = 13 And Centavos <> 0
           Bilhao = Mid$(ValChr, 1, 1)
           Milhao = Val(Mid$(ValChr, 2, 3))
           Mil = Val(Mid$(ValChr, 5, 3))
           Unid = Val(Mid$(ValChr, 8, 3))
       
       Case Is = 12 And Centavos <> 0
           Milhao = Val(Mid$(ValChr, 1, 3))
           Mil = Val(Mid$(ValChr, 4, 3))
           Unid = Val(Mid$(ValChr, 7, 3))
       
       Case Is = 11 And Centavos <> 0
           Milhao = Val(Mid$(ValChr, 1, 2))
           Mil = Val(Mid$(ValChr, 3, 3))
           Unid = Val(Mid$(ValChr, 6, 3))
       
       Case Is = 10 And Centavos <> 0
           Milhao = Val(Mid$(ValChr, 1, 1))
           Mil = Val(Mid$(ValChr, 2, 3))
           Unid = Val(Mid$(ValChr, 5, 3))

       Case Is = 9 And Centavos <> 0
           Mil = Val(Mid$(ValChr, 1, 3))
           Unid = Val(Mid$(ValChr, 4, 3))
       
       Case Is = 8 And Centavos <> 0
           Mil = Val(Mid$(ValChr, 1, 2))
           Unid = Val(Mid$(ValChr, 3, 3))
       
       Case Is = 7 And Centavos <> 0
           Mil = Val(Mid$(ValChr, 1, 1))
           Unid = Val(Mid$(ValChr, 2, 3))
       
       Case Is = 6 And Centavos <> 0
           Unid = Val(Mid$(ValChr, 1, 3))
       Case Is = 5 And Centavos <> 0
           Unid = Val(Mid$(ValChr, 1, 2))
       Case Is = 4 And Centavos <> 0
           Unid = Val(Mid$(ValChr, 1, 1))
       
       Case Is = 15 And Centavos = 0
           Trilhao = Val(Left$(ValChr, 3))
           Bilhao = Mid$(ValChr, 4, 3)
           Milhao = Val(Mid$(ValChr, 7, 3))
           Mil = Val(Mid$(ValChr, 10, 3))
           Unid = Val(Mid$(ValChr, 13, 3))
       Case Is = 14 And Centavos = 0
           Trilhao = Val(Left$(ValChr, 2))
           Bilhao = Mid$(ValChr, 3, 3)
           Milhao = Val(Mid$(ValChr, 6, 3))
           Mil = Val(Mid$(ValChr, 9, 3))
           Unid = Val(Mid$(ValChr, 12, 3))

       Case Is = 13 And Centavos = 0
           Trilhao = Val(Left$(ValChr, 1))
           Bilhao = Mid$(ValChr, 2, 3)
           Milhao = Val(Mid$(ValChr, 5, 3))
           Mil = Val(Mid$(ValChr, 9, 3))
           Unid = Val(Mid$(ValChr, 11, 3))

       Case Is = 12 And Centavos = 0
           Bilhao = Mid$(ValChr, 1, 3)
           Milhao = Val(Mid$(ValChr, 4, 3))
           Mil = Val(Mid$(ValChr, 7, 3))
           Unid = Val(Mid$(ValChr, 10, 3))
       
       Case Is = 11 And Centavos = 0
           Bilhao = Mid$(ValChr, 1, 2)
           Milhao = Val(Mid$(ValChr, 3, 3))
           Mil = Val(Mid$(ValChr, 6, 3))
           Unid = Val(Mid$(ValChr, 9, 3))
       
       Case Is = 10 And Centavos = 0
           Bilhao = Mid$(ValChr, 1, 1)
           Milhao = Val(Mid$(ValChr, 2, 3))
           Mil = Val(Mid$(ValChr, 5, 3))
           Unid = Val(Mid$(ValChr, 8, 3))

       Case Is = 9 And Centavos = 0
           Milhao = Val(Mid$(ValChr, 1, 3))
           Mil = Val(Mid$(ValChr, 4, 3))
           Unid = Val(Mid$(ValChr, 7, 3))
       
       Case Is = 8 And Centavos = 0
           Milhao = Val(Mid$(ValChr, 1, 2))
           Mil = Val(Mid$(ValChr, 3, 3))
           Unid = Val(Mid$(ValChr, 6, 3))
       
       Case Is = 7 And Centavos = 0
           Milhao = Val(Mid$(ValChr, 1, 1))
           Mil = Val(Mid$(ValChr, 2, 3))
           Unid = Val(Mid$(ValChr, 5, 3))

       Case Is = 6 And Centavos = 0
           Mil = Val(Mid$(ValChr, 1, 3))
           Unid = Val(Mid$(ValChr, 4, 3))
       Case Is = 5 And Centavos = 0
           Mil = Val(Mid$(ValChr, 1, 2))
           Unid = Val(Mid$(ValChr, 3, 3))
       Case Is = 4 And Centavos = 0
           Mil = Val(Mid$(ValChr, 1, 1))
           Unid = Val(Mid$(ValChr, 2, 3))

       Case Is = 3 And Centavos = 0
           Unid = Val(Mid$(ValChr, 1, 3))
       Case Is = 2 And Centavos = 0
           Unid = Val(Mid$(ValChr, 1, 2))
       Case Is = 1 And Centavos = 0
           Unid = Val(Mid$(ValChr, 1, 1))


End Select

Ext = ""
If Trilhao > 0 Then
 Ext = IIf(Trilhao = 1, "Hum", ExtQtd(Trilhao))
 Ext = Ext + IIf(Trilhao = 1, " Trilhão", " Trilhões")
 If Bilhao = 0 And Milhao = 0 And Mil = 0 And Unid = 0 Then
  Ext = Ext + " de"
 Else
  Ext = Ext + ","
 End If
End If
If Bilhao > 0 Then
 Ext = Ext + " " + IIf(Bilhao = 1 And Valor < 2000000000, "Hum", ExtQtd(Bilhao))
 Ext = Ext + IIf(Bilhao = 1, " Bilhão", " Bilhões")
 If Milhao = 0 And Mil = 0 And Unid = 0 Then
  Ext = Ext + " de"
 Else
  Ext = Ext + ","
 End If
End If
If Milhao > 0 Then
 Ext = Ext + " " + IIf(Milhao = 1 And Valor < 2000000, "Hum", ExtQtd(Milhao))
 Ext = Ext + IIf(Milhao = 1, " Milhão", " Milhões")
 If Mil = 0 And Unid = 0 Then
  Ext = Ext + " de"
 Else
  Ext = Ext + ","
 End If
End If
If Mil > 0 Then
 Ext = Ext + " " + IIf(Mil = 1 And Valor < 2000, "Hum", ExtQtd(Mil))
 Ext = Ext + " mil"
End If
If Unid > 0 Then
 If Valor > 999.99 Then
 
  Ext = Ext + IIf(Mil = 0 And Unid > 100, " ,", " e")
 End If
 Ext = Ext + " " + IIf(Unid = 1 And Valor < 2, "Hum", ExtQtd(Unid))
End If
If Valor >= 1 Then

 Ext = Ext + " " + Trim$(IIf(Int(Valor) = 1, NmMoe_1sg, NmMoe_1pl))
End If
If Centavos > 0 Then
 If Valor >= 1 Then
  Ext = Ext + " e"
 End If
 Ext = Ext + " " + ExtQtd(Centavos)
 Ext = Ext + " " + Trim$(IIf(Centavos = 1, NmMoe_2sg, NmMoe_2pl))
End If
Ext = Trim$(Ext)
Extenso = Ext

erro:
        Select Case Err
               Case Is = 13
                    MsgBox "Tipos incompatíveis de dados" + Chr(13) + "Valor deve ser numérico", 16, "Erro nos dados"
               Case Is = 6
                    MsgBox "Talvez não seja possivel realizar o extenso deste número." + Chr(13) + Chr(13) + "Por favor,verifique o extenso !", 16, "Estouro de capacidade"
        End Select
        Exit Function
End Function

Function ExtQtd(Valor)

Dim Ext, Centena, Dezena, Unidade, ValChr

ValChr = CCur(Valor)


If Len(Valor) = 3 Then

    Centena = Val(Left$(ValChr, 1))

    Dezena = Val(Mid$(ValChr, 2, 1))

    Unidade = Val(Right$(ValChr, 1))
Else
    Select Case Len(Valor)
           Case Is = 2
                Dezena = Val(Left$(ValChr, 1))
                Unidade = Val(Right$(ValChr, 1))
           Case Is = 1
                Unidade = Val(Right$(ValChr, 1))
    End Select
End If

Ext = ""
If Centena > 0 Then
 Select Case Centena <> 0
        Case Centena = 1
              If Dezena = 0 And Unidade = 0 Then
                   Ext = Ext + "Cem"
              Else
                   Ext = Ext + "Cento"
              End If
         Case Centena = 2
              Ext = Ext + "Duzentos"
         Case Centena = 3
              Ext = Ext + "Trezentos"
         Case Centena = 4
              Ext = Ext + "Quatrocentos"
         Case Centena = 5
              Ext = Ext + "Quinhentos"
         Case Centena = 6
              Ext = Ext + "Seiscentos"
         Case Centena = 7
              Ext = Ext + "Setecentos"
         Case Centena = 8
              Ext = Ext + "Oitocentos"
         Case Centena = 9
              Ext = Ext + "Novecentos"
 End Select
 If Dezena <> 0 Or Unidade <> 0 Then
  Ext = Ext + " e "
 End If
End If
If Dezena > 0 Then
 Select Case Dezena > 0
 Case Dezena = 1
  Select Case Unidade <> 0 Or Unidade = 0
  Case Unidade = 0
   Ext = Ext + "Dez"
  Case Unidade = 1
   Ext = Ext + "Onze"
  Case Unidade = 2
   Ext = Ext + "Doze"
  Case Unidade = 3
   Ext = Ext + "Treze"
  Case Unidade = 4
   Ext = Ext + "Quatorze"
  Case Unidade = 5
   Ext = Ext + "Quinze"
  Case Unidade = 6
   Ext = Ext + "Dezesseis"
  Case Unidade = 7
   Ext = Ext + "Dezessete"
  Case Unidade = 8
   Ext = Ext + "Dezoito"
  Case Unidade = 9
   Ext = Ext + "Dezenove"
  End Select
 Case Dezena = 2
  Ext = Ext + "Vinte"
 Case Dezena = 3
  Ext = Ext + "Trinta"
 Case Dezena = 4
  Ext = Ext + "Quarenta"
 Case Dezena = 5
  Ext = Ext + "Cinqüenta"
 Case Dezena = 6
  Ext = Ext + "Sessenta"
 Case Dezena = 7
  Ext = Ext + "Setenta"
 Case Dezena = 8
  Ext = Ext + "Oitenta"
 Case Dezena = 9
  Ext = Ext + "Noventa"
 End Select
 If Unidade <> 0 And Dezena > 1 Then
  Ext = Ext + " e "
 End If
End If
If Unidade > 0 And Dezena <> 1 Then
 Select Case Unidade > 0
 Case Unidade = 1
  Ext = Ext + "Um"
 Case Unidade = 2
  Ext = Ext + "Dois"
 Case Unidade = 3
  Ext = Ext + "Três"
 Case Unidade = 4
  Ext = Ext + "Quatro"
 Case Unidade = 5
  Ext = Ext + "Cinco"
 Case Unidade = 6
  Ext = Ext + "Seis"
 Case Unidade = 7
  Ext = Ext + "Sete"
 Case Unidade = 8
  Ext = Ext + "Oito"
 Case Unidade = 9
  Ext = Ext + "Nove"
 End Select
End If
ExtQtd = Ext
End Function



'Validação de CNPJ

Public Function CNPJValido(CNPJ As String) As Boolean
 Dim A As Integer
 Dim J As Integer
 Dim i As Integer
 Dim D1 As Integer
 Dim D2 As Integer
 
  If Len(CNPJ) = 8 And Val(CNPJ) > 0 Then
     A = 0
     J = 0
     D1 = 0
     For i = 1 To 7
      A = Val(Mid(CNPJ, i, 1))
                If (i Mod 2) <> 0 Then A = A * 2
                If A > 9 Then
                   J = J + Int(A / 10) + (A Mod 10)
                Else
                   J = J + A
                End If
     Next i
     D1 = IIf((J Mod 10) <> 0, 10 - (J Mod 10), 0)
                If D1 = Val(Mid(CNPJ, 8, 1)) Then
                   CNPJValido = True
                Else
                   CNPJValido = False
                End If
  Else
            If Len(CNPJ) = 14 And Val(CNPJ) > 0 Then
               A = 0
               i = 0
               D1 = 0
               D2 = 0
               J = 5
               For i = 1 To 12 Step 1
                   A = A + (Val(Mid(CNPJ, i, 1)) * J)
                   J = IIf(J > 2, J - 1, 9)
               Next i
               A = A Mod 11
               D1 = IIf(A > 1, 11 - A, 0)
               A = 0
               i = 0
               J = 6
               For i = 1 To 13 Step 1
                   A = A + (Val(Mid(CNPJ, i, 1)) * J)
                   J = IIf(J > 2, J - 1, 9)
               Next i
               A = A Mod 11
               D2 = IIf(A > 1, 11 - A, 0)
                           If (D1 = Val(Mid(CNPJ, 13, 1)) And D2 = Val(Mid(CNPJ, 14, 1))) Then
                              CNPJValido = True
                           Else
                              CNPJValido = False
                           End If
            Else
               CNPJValido = False
            End If
  End If
End Function

'Validação de Data

Public Function ValidaData(DataInformada As String) As Boolean
On Error GoTo falha
If DataInformada <> "__/__/__" Then
    ValidaData = False
    'variáveis necessárias
    Dim dia As Integer
    Dim mes As Integer
    Dim resto As String
    Dim ANO As Integer
    Dim calc As Integer
    
    'extraindo o dia mês e ano da data digitada
    dia = Left(DataInformada, InStr(DataInformada, "/") - 1)
    resto = Mid(DataInformada, InStr(DataInformada, "/") + 1)
    mes = Left(resto, InStr(resto, "/") - 1)
    ANO = Right(resto, 2)
    
    'verifica se todos os dígitos foram preenchidos
    If InStr(dia, "_") <> False Or InStr(mes, "_") <> False Or InStr(ANO, "_") <> False Then
        Exit Function
    End If
    calc = ANO Mod 4

    'para anos bissextos
    If calc = 0 Then
        If mes > 12 Then
            Exit Function
        End If
        If dia > 31 Then
            Exit Function
        End If
        If dia > 29 And mes = 2 Then
            Exit Function
        End If
        If dia > 30 And mes = 4 Then
            Exit Function
        End If
        If dia > 30 And mes = 6 Then
            Exit Function
        End If
        If dia > 30 And mes = 9 Then
            Exit Function
        End If
        If dia > 30 And mes = 11 Then
            Exit Function
        End If
    End If
    
    'para anos não bissextos
    If calc <> 0 Then
        If mes > 12 Then
            Exit Function
        End If
        If dia > 31 Then
            Exit Function
        End If
        If dia > 28 And mes = 2 Then
            Exit Function
        End If
        If dia > 30 And mes = 4 Then
            Exit Function
        End If
        If dia > 30 And mes = 6 Then
            Exit Function
        End If
        If dia > 30 And mes = 9 Then
            Exit Function
        End If
        If dia > 30 And mes = 11 Then
            Exit Function
        End If
    End If
End If
ValidaData = True
falha:
End Function

Public Function VerifCPF(CPF As String) As Boolean
    CPF = Mid(CPF, 1, 3) & Mid(CPF, 5, 3) & Mid(CPF, 9, 3) & Mid(CPF, 13, 2)
    
    Dim MULT1 As Integer
    Dim MULT2 As Integer
    Dim DIG1 As Integer
    Dim DIG2 As Integer
    Dim Y1 As String
    Dim Y2 As String
    Dim Z2 As String

    Let Y1 = 9
    Let Y2 = 10
    Let Z2 = 11
    MULT1 = 10
    MULT2 = 11

    If Len(CPF) = 10 Then
        Let MULT1 = 9
        Let MULT2 = 10
        Let Z2 = 10
        Let Y1 = 8
        Let Y2 = 9
    End If

    For CONT = 1 To Y1
        DIG1 = DIG1 + (Val(Mid$(CPF, CONT, 1)) * MULT1)
        MULT1 = MULT1 - 1
    Next

    For CONT = 1 To Y2
        DIG2 = DIG2 + (Val(Mid$(CPF, CONT, 1)) * MULT2)
        MULT2 = MULT2 - 1
    Next

    DIG1 = (DIG1 * 10) Mod 11
    DIG2 = (DIG2 * 10) Mod 11
    If DIG1 = 10 Then DIG1 = 0
    If DIG2 = 10 Then DIG2 = 0

    VerifCPF = True

    If DIG1 <> Mid$(CPF, Y2, 1) Then VerifCPF = False
    If DIG2 <> Mid$(CPF, Z2, 1) Then VerifCPF = False

End Function


