Set objUser = WScript.CreateObject("WScript.Network")

wuser=objUser.UserName
If Time <= "12:00:00" Then
MsgBox ("Bom dia, "Wuser+"! Aviso completo")
ElseIf Time >= "12:00:01" And Time <= "18:00:00" Then

MsgBox ("Boa tarde, "Wuser+"! Aviso completo")
Else
MsgBox ("Boa noite, ")

End If
Wscript.Quit