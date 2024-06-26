' Definir variáveis
Const LOWCASE = "abcdefghijklmnopqrstuvwxyz"
Const UPPERCASE = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const NUMBERS = "0123456789"
Const SYMBOLS = "!@#$%^&*()_-+={}[]|;':<>,.?"
Const PASSWORD_LENGTH = 5 ' Altere o tamanho da senha aqui

' Criar arrays
arrLowerCase = Split(LOWCASE)
arrUpperCase = Split(UPPERCASE)
arrNumbers = Split(NUMBERS)
arrSymbols = Split(SYMBOLS)

' Função para gerar caractere aleatório
Function GetRandomChar(arrChars)
    Dim randomIndex
    randomIndex = Int(Rnd() * UBound(arrChars))
    GetRandomChar = arrChars(randomIndex)
End Function

' Criar string vazia para armazenar a senha
strPassword = ""

' Gerar senha com base na complexidade desejada
For i = 1 To PASSWORD_LENGTH
    ' Nível de complexidade baixo (apenas letras minúsculas)
    If PASSWORD_LENGTH <= 8 Then
        strPassword = strPassword & GetRandomChar(arrLowerCase)
    ' Nível de complexidade médio (letras minúsculas e maiúsculas)
    ElseIf PASSWORD_LENGTH <= 12 Then
        strPassword = strPassword & GetRandomChar(arrLowerCase) & GetRandomChar(arrUpperCase)
    ' Nível de complexidade alto (letras minúsculas, maiúsculas, números e símbolos)
    Else
        strPassword = strPassword & GetRandomChar(arrLowerCase) & GetRandomChar(arrUpperCase) & GetRandomChar(arrNumbers) & GetRandomChar(arrSymbols)
    End If
Next

' Exibir a senha gerada
MsgBox "Sua senha aleatoria e: " & strPassword
