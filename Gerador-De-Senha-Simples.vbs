' Definir variáveis
strMasterPassword = "senhaMestra123" ' Defina sua senha mestra aqui
strFileName = "minhas_senhas.txt" ' Nome do arquivo de senhas

' Função para criptografar texto
Function EncryptText(strText)
    strEncryptedText = Base64Encode(strText) ' Criptografia base64 simples
    EncryptText = strEncryptedText
End Function

' Função para descriptografar texto
Function DecryptText(strEncryptedText)
    strDecryptedText = Base64Decode(strEncryptedText) ' Descriptografia base64 simples
    DecryptText = strDecryptedText
End Function

' Função para adicionar uma nova senha
Sub AddNewPassword()
    ' Obter nome de usuário e senha do usuário
    strUsername = InputBox("Digite o nome de usuário: ")
    strPassword = InputBox("Digite a senha: ")

    ' Verificar se a senha mestra está correta
    strMasterPasswordInput = InputBox("Digite a senha mestra: ")
    If strMasterPasswordInput <> strMasterPassword Then
        MsgBox "Senha mestra incorreta. Acesso negado."
        Exit Sub
    End If

    ' Criptografar nome de usuário e senha
    strEncryptedUsername = EncryptText(strUsername)
    strEncryptedPassword = EncryptText(strPassword)

    ' Ler conteúdo do arquivo de senhas (se existir)
    If fso.FileExists(strFileName) Then
        Set f = fso.OpenTextFile(strFileName, 1) ' 1 para leitura
        strFileContent = f.ReadAll
        f.Close
    Else
        strFileContent = ""
    End If

    ' Adicionar nova entrada no conteúdo do arquivo
    strNewEntry = strEncryptedUsername & "|" & strEncryptedPassword & vbCrLf
    strFileContent = strFileContent & strNewEntry

    ' Gravar conteúdo atualizado no arquivo de senhas
    Set f = fso.CreateTextFile(strFileName, 2) ' 2 para gravação
    f.Write strFileContent
    f.Close

    ' Exibir mensagem de sucesso
    MsgBox "Senha adicionada com sucesso!"
End Sub

' Função para exibir todas as senhas
Sub ShowAllPasswords()
    ' Verificar se a senha mestra está correta
    strMasterPasswordInput = InputBox("Digite a senha mestra: ")
    If strMasterPasswordInput <> strMasterPassword Then
        MsgBox "Senha mestra incorreta. Acesso negado."
        Exit Sub
    End If

    ' Ler conteúdo do arquivo de
End Sub