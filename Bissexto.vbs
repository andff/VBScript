End Select

' Função para verificar se um ano é bissexto
Function LeapYear(intYear)
    If intYear Mod 4 = 0 Then
        If intYear Mod 100 <> 0 Or intYear Mod 400 = 0 Then
            LeapYear = True
        Else
            LeapYear = False
        End If
    Else
        LeapYear = False
    End If
End Function

' Exibir mensagem de data válida
MsgBox "Data válida: " & strDate
