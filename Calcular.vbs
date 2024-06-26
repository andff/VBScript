' Obter números do usuário
num1 = InputBox("Digite o primeiro número: ")
num2 = InputBox("Digite o segundo número: ")

' Obter operação do usuário
strOperation = InputBox("Digite a operacao (+, -, *, /): ")

' Realizar operação
Select Case strOperation
    Case "+"
        result = num1 + num2
    Case "-"
        result = num1 - num2
    Case "*"
        result = num1 * num2
    Case "/"
        If num2 = 0 Then
            MsgBox "Erro: Divisao por zero."
            'Exit Sub
        End If
        result = num1 / num2
    Default
        MsgBox "Operacao invalida."
        'Exit Sub
End Select

' Exibir resultado
MsgBox "Resultado: " & result
