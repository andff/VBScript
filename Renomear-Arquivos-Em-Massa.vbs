' Obter diretório do usuário
strFolder = InputBox("Digite o caminho do diretório: ")

' Obter nome base e sufixo do usuário
strBaseName = InputBox("Digite o nome base para os arquivos: ")
strSuffix = InputBox("Digite o sufixo para os arquivos: ")

' Renomear arquivos
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(strFolder)

For Each file In folder.Files
    strNewName = strBaseName & "_" & file.Name & strSuffix
    file.Name = strNewName
Next

' Exibir mensagem de sucesso
MsgBox "Arquivos renomeados com sucesso!"
