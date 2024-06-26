' Definir número secreto
intSecretNumber = Int(Rnd() * 100) + 1

' Variáveis para contagem de tentativas e adivinhação
intAttempts = 0
intGuess = 0

' Loop do jogo
Do Until intGuess = intSecretNumber
    intAttempts = intAttempts + 1
    intGuess = InputBox("Adivinhe o número secreto (1-100): ")

    If intGuess > intSecretNumber Then
        MsgBox "Seu palpite foi muito alto!"
    ElseIf intGuess < intSecretNumber Then
        MsgBox "Seu palpite foi muito baixo!"
    End If
Loop

' Exibir mensagem de vitória
MsgBox "Parabéns! Você adivinhou o número em " & intAttempts & " tentativas."
