Function ValidaCPF(CPF As String) As Boolean
    Dim i As Integer
    Dim d1 As Integer, d2 As Integer
    Dim digit1 As Integer, digit2 As Integer
    
    ' Remove caracteres que não são números
    CPF = Replace(CPF, ".", "")
    CPF = Replace(CPF, "-", "")
    
    ' Verifica se o CPF tem 11 dígitos
    If Len(CPF) <> 11 Then
        ValidaCPF = False
        Exit Function
    End If
    
    ' Verifica se todos os dígitos são iguais
    If CPF = String(11, Left(CPF, 1)) Then
        ValidaCPF = False
        Exit Function
    End If
    
    ' Calcula o primeiro dígito verificador
    For i = 1 To 9
        d1 = d1 + Mid(CPF, i, 1) * (11 - i)
    Next i
    digit1 = 11 - (d1 Mod 11)
    If digit1 >= 10 Then digit1 = 0
    
    ' Calcula o segundo dígito verificador
    For i = 1 To 10
        d2 = d2 + Mid(CPF, i, 1) * (12 - i)
    Next i
    digit2 = 11 - (d2 Mod 11)
    If digit2 >= 10 Then digit2 = 0
    
    ' Verifica se os dígitos calculados correspondem aos do CPF
    If digit1 = Mid(CPF, 10, 1) And digit2 = Mid(CPF, 11, 1) Then
        ValidaCPF = True
    Else
        ValidaCPF = False
    End If
End Function
