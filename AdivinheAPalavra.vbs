' Define variables
Const MAX_GUESSES = 6 ' Maximum allowed guesses
strSecretWord = "secreto" ' Replace with your secret word
arrGuessedLetters = Array() ' Array to store guessed letters
strDisplayWord = "" ' Initialize word with underscores

' Initialize display word
For i = 1 To Len(strSecretWord)
    strDisplayWord = strDisplayWord & "_"
Next

' Game loop
intGuessesLeft = MAX_GUESSES
Do Until intGuessesLeft = 0 Or strDisplayWord = strSecretWord
    ' Display current word and guesses remaining
    MsgBox "Palavra: " & strDisplayWord & vbCrLf & "Letras adivinhadas: " & Join(arrGuessedLetters, ", ") & vbCrLf & "Adivinhas restantes: " & intGuessesLeft

    ' Get player's guess
    strGuess = UCase(InputBox("Digite uma letra: "))

    ' Check if letter has already been guessed
    If IsInStr(Join(arrGuessedLetters, ", "), strGuess) Then
        MsgBox "Você já adivinhou essa letra!"
        Continue Do
    End If

    ' Add guess to the array
    arrGuessedLetters(UBound(arrGuessedLetters) + 1) = strGuess

    ' Check if guess is correct
    If InStr(strSecretWord, strGuess) = 0 Then
        ' Incorrect guess
        MsgBox "Letra incorreta!"
        intGuessesLeft = intGuessesLeft - 1
    Else
        ' Correct guess
        ' Update display word
        For i = 1 To Len(strSecretWord)
            If Mid(strSecretWord, i, 1) = strGuess Then
                strDisplayWord = Left(strDisplayWord, i - 1) & strGuess & Right(strDisplayWord, Len(strSecretWord) - i)
            End If
        Next

        ' Check if word is complete
        If strDisplayWord = strSecretWord Then
            MsgBox "Parabéns! Você adivinhou a palavra!"
            Exit Do
        End If
    End If
Loop

' If guesses run out, reveal the word
If intGuessesLeft = 0 Then
    MsgBox "Você perdeu! A palavra era: " & strSecretWord
End If
