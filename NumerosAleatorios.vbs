' Obter intervalo do usuário
intMin = InputBox("Digite o limite mínimo: ")
intMax = InputBox("Digite o limite máximo: ")

' Validar intervalo
If intMin > intMax Then
    MsgBox "O limite mínimo não pode ser maior que o limite máximo."
    Exit Sub
End If

' Gerar número aleatório
intRandom = Int(Rnd() * (intMax - intMin + 1)) + intMin

' Exibir número aleatório
MsgBox "Número aleatório: " & intRandom
