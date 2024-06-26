' Set remote computer name and command
strRemoteComputer = "NOME_DO_COMPUTADOR_REMOTO"
strCommand = "ipconfig" ' Replace with your command

' Create WMI connection
Set objWMI = CreateObject("WbemScripting.SWbemLocator")
objWMI.ConnectServer strRemoteComputer, "root\cimv2", "", "", 0, 0, 0, 0

' Get the process namespace
Set objExec = objWMI.GetNamespace("root\cimv2")

' Execute the command and get the output
Set objProcess = objExec.ExecMethod_("ExecMethod", Array(strCommand))
strOutput = objProcess.StdOutText

' Display the output
MsgBox "Saída do comando em " & strRemoteComputer & ":" & vbCrLf & strOutput
