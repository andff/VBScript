' Set folder path and share name
strFolderPath = "C:\caminho\para\sua\pasta"
strShareName = "NOME_DO_SEU_COMPARTILHAMENTO"

' Create the network share
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strFolderPath)
objFolder.Share(strShareName)

' Display confirmation message
MsgBox "Pasta " & strFolderPath & " compartilhada como " & strShareName
