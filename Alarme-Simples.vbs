' Criar objeto WScript.Network
Set objNet = CreateObject("WScript.Network")

' Definir a frequência e a duração do som
intFrequency = 500 ' Frequência em Hz (ajuste para alterar o tom)
intDuration = 1000 ' Duração em milissegundos (ajuste para alterar o tempo do som)

' Reproduzir o som do alarme
objNet.Beep(intFrequency, intDuration)

' Opcional: Exibir mensagem de aviso
MsgBox "Alarme ativado!"
