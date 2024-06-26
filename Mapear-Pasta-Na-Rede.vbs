' Set network drive letter, server name, and share name
strDriveLetter = "M:" ' Replace with your desired drive letter
strServerName = "labservidor" ' Replace with the server name
strShareName = "ti$" ' Replace with the share name

' Map the network drive
Set objWSH = CreateObject("WScript.Shell")
objWSH.Run "NET USE " & strDriveLetter: & "\\" & strServerName & "\" & strShareName & " /PERSISTENT", 0 ' 0 = Hide window

' Display confirmation message
MsgBox "Unidade de rede " & strDriveLetter & " mapeada para " & strServerName & "\" & strShareName & " com sucesso!"
