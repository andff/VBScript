' Get the current user's name
strUserName = WScript.CreateObject("WScript.Shell").Environment("UserName")

' Get the current date and time
strDateTime = Now()

' Display the greeting message
MsgBox "Olá, " & strUserName & "! Seja bem-vindo ao sistema." & vbCrLf & _
       "Data de login: " & FormatDateTime(strDateTime, vbShortDate) & " " & _
       FormatDateTime(strDateTime, vbLongTime)

' Open the log file in append mode
Set objLogFile = CreateObject("Scripting.FileSystemObject")
Set objLog = objLogFile.OpenTextFile("C:\Users\" & strUserName & "\log.txt", True)

' Write the user's name, date, and time to the log file
objLog.WriteLine strUserName & "," & FormatDateTime(strDateTime, vbLongDate) & "," & _
               FormatDateTime(strDateTime, vbLongTime)

' Close the log file
objLog.Close

' Display a success message
MsgBox "Seu login foi registrado no arquivo de log."
