'----------------------------------------------------------------------
'    pubprn.vbs - publish printers from a non Windows 2000 server into Windows 2000 DS
'    
'
'     Arguments are:-
'        server - format server
'        DS container - format "LDAP:\\CN=...,DC=...."
'
'
'    Copyright (c) Microsoft Corporation 1997
'   All Rights Reserved
'----------------------------------------------------------------------

'--- Begin Error Strings ---

Dim L_PubprnUsage1_text
Dim L_PubprnUsage2_text
Dim L_PubprnUsage3_text      
Dim L_PubprnUsage4_text      
Dim L_PubprnUsage5_text      
Dim L_PubprnUsage6_text      

Dim L_GetObjectError1_text
Dim L_GetObjectError2_text

Dim L_PublishError1_text
Dim L_PublishError2_text     
Dim L_PublishError3_text
Dim L_PublishSuccess1_text


L_PubprnUsage1_text      =   "Uso: [cscript] pubprn.vbs servidor ""LDAP://OU=..,DC=..."""
L_PubprnUsage2_text      =   "       servidor é um nome de servidor do Windows (por exemplo: servidor) ou nome UNC de impressora (\\servidor\impressora)"
L_PubprnUsage3_text      =   "       ""LDAP://CN=...,DC=..."" é o caminho DS do contêiner de destino"
L_PubprnUsage4_text      =   ""
L_PubprnUsage5_text      =   "Exemplo 1: pubprn.vbs Meu_Servidor ""LDAP://CN=Meu_Recipiente,DC=Meu_Domínio,DC=Empresa,DC=Com"""
L_PubprnUsage6_text      =   "Exemplo 2: pubprn.vbs \\Meu_servidor\Impressora ""LDAP://CN=Meu_Recipiente,DC=Meu_Domínio,DC=Empresa,DC=Com"""

L_GetObjectError1_text   =   "Erro: caminho "
L_GetObjectError2_text   =   " não encontrado."
L_GetObjectError3_text   =   "Erro: não é possível acessar "

L_PublishError1_text     =   "Erro: Pubprn não pode publicar impressoras de "
L_PublishError2_text     =   " porque está executando o Windows 2000 ou posterior."
L_PublishError3_text     =   "Falha ao publicar impressora "
L_PublishError4_text     =   "Erro: "
L_PublishSuccess1_text   =   "Impressora publicada: "

'--- End Error Strings ---


set Args = Wscript.Arguments
if args.count < 2 then
    wscript.echo L_PubprnUsage1_text
    wscript.echo L_PubprnUsage2_text
    wscript.echo L_PubprnUsage3_text
    wscript.echo L_PubprnUsage4_text
    wscript.echo L_PubprnUsage5_text
    wscript.echo L_PubprnUsage6_text
    wscript.quit(1)
end if

ServerName= args(0)
Container = args(1)

if 1 <> InStr(1, Container, "LDAP://", vbTextCompare) then
    wscript.echo L_GetObjectError1_text & Container & L_GetObjectError2_text
    wscript.quit(1)
end if

on error resume next
Set PQContainer = GetObject(Container)

if err then
    wscript.echo L_GetObjectError1_text & Container & L_GetObjectError2_text
    wscript.quit(1)
end if
on error goto 0



if left(ServerName,1) = "\" then

    PublishPrinter ServerName, ServerName, Container

else

    on error resume next

    Set PrintServer = GetObject("WinNT://" & ServerName & ",computer")

    if err then
        wscript.echo L_GetObjectError3_text & ServerName & ": " & err.Description
        wscript.quit(1)
    end if

    on error goto 0


    For Each Printer In PrintServer
        if Printer.class = "PrintQueue" then PublishPrinter Printer.PrinterPath, ServerName, Container
    Next


end if




sub PublishPrinter(UNC, ServerName, Container)

    
    Set PQ = WScript.CreateObject("OlePrn.DSPrintQueue.1")

    PQ.UNCName = UNC
    PQ.Container = Container

    on error resume next

    PQ.Publish(2)

    if err then
        if err.number = -2147024772 then
            wscript.echo L_PublishError1_text & Chr(34) & ServerName & Chr(34) & L_PublishError2_text
            wscript.quit(1)
        else
            wscript.echo L_PublishError3_text & Chr(34) & UNC & Chr(34) & "."
            wscript.echo L_PublishError4_text & err.Description
        end if
    else
        wscript.echo L_PublishSuccess1_text & PQ.Path
    end if

    Set PQ = nothing

end sub
