' Set message to send
strMessage = "Olá a todos!"

' Create the NET SEND command
strCommand = "NET SEND * " & strMessage

' Execute the NET SEND command
Set objShell = CreateObject("WScript.Shell")
objShell.Run strCommand, 0 ' 0 = Hide window

' Display confirmation message
MsgBox "Mensagem de rede enviada com sucesso!"
