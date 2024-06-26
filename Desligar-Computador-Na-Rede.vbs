' Definir variáveis
strComputerName = "NOME_DO_COMPUTADOR_REMOTO" ' Substitua pelo nome do computador remoto
strUsername = "NOME_DO_USUARIO" ' Substitua pelo nome de usuário com privilégios administrativos no computador remoto
strPassword = "SENHA_DO_USUARIO" ' Substitua pela senha do usuário remoto

' Criar conexão WMI
Set objWMI = CreateObject("WbemScripting.SWbemLocator")
objWMI.ConnectServer strComputerName, "root\cimv2", strUsername, strPassword, 0, 0, 0, 0

' Obter namespace do processo
Set objExec = objWMI.GetNamespace("root\cimv2")

' Executar comando de desligamento (força o desligamento)
objExec.ExecMethod_("ExecMethod", Array("shutdown /s /f /t 0"))

' Exibir mensagem de confirmação
MsgBox "O computador " & strComputerName & " foi desligado."
