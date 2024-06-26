' Obter data do usuário
strDate = InputBox("Digite uma data no formato dd/mm/aaaa: ")

' Dividir a data em partes
arrDate = Split(strDate, "/")

' Validar formato da data
If UBound(arrDate) <> 2 Then
    MsgBox "Formato de data inválido. Use dd/mm/aaaa."
    Exit Sub
End If

' Converter partes da data para números
intDay = Val(arrDate(0))
intMonth = Val(arrDate(1))
intYear = Val(arrDate(2))

' Validar dia
If intDay <= 0 Or intDay > 31 Then
    MsgBox "Dia inválido. Digite um valor entre 1 e 31."
    Exit Sub
End If

' Validar mês
If intMonth <= 0 Or intMonth > 12 Then
    MsgBox "Mês inválido. Digite um valor entre 1 e 12."
    Exit Sub
End If

' Validar ano
If intYear < 1000 Or intYear > 9999 Then
    MsgBox "Ano inválido. Digite um valor entre 1000 e 9999."
    Exit Sub
End If

' Validar data com base em dias do mês
Select Case intMonth
    Case 2, 4, 6, 9, 11
        If intDay > 29 Then
            MsgBox "Data inválida. Este mês só tem 29 dias."
            Exit Sub
        End If
    Case 4
        If intYear Is LeapYear(intYear) And intDay > 29 Then
            MsgBox "Data inválida. Este ano bissexto só tem 29 dias em fevereiro."
            Exit Sub
        End
