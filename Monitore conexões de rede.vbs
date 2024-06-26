' Create WMI connection
Set objWMI = CreateObject("WbemScripting.SWbemLocator")
objWMI.ConnectServer ".", "root\cimv2", "", "", 0, 0, 0, 0

' Get the network connection namespace
Set objNetCon = objWMI.GetNamespace("root\cimv2")

' Query network connections
Set objQuery = objNetCon.ExecQuery("SELECT * FROM Win32_NetworkConnection")

' Display connection information
MsgBox "Conexões de rede ativas
