' Set connection parameters
strDomainName = "SEU_DOMINIO" ' Replace with your domain name
strUsername = "SEU_USUARIO" ' Replace with your domain username
strPassword = "SEU_PASSWORD" ' Replace with your domain password

' Establish connection to Active Directory
Set objAD = CreateObject("ADODB.Connection")
strADConnectionString = "ADSProvider=LDAP;SERVER=" & strDomainName & ";DOMAIN=" & strDomainName
objAD.ConnectionString = strADConnectionString
objAD.Open

' Set search base
strSearchBase = "LDAP://CN=Users," & strDomainName

' Create search query
strSearchQuery = "(objectClass=user)"

' Execute search and retrieve results
Set objSearcher = CreateObject("ADODB.Command")
objSearcher.ActiveConnection = objAD
objSearcher.CommandText = "SELECT name, mail FROM " & strSearchBase & " WHERE " & strSearchQuery
Set objRS = objSearcher.Execute

' Display user information
MsgBox "Lista de usuários do Active Directory:"
Do Until objRS.EOF
    strName = objRS("name")
    strEmail = objRS("mail")
    MsgBox strName & ": " & strEmail
    objRS.MoveNext
Loop

objRS.Close
objSearcher.ActiveConnection = Nothing
objSearcher.Set Null
objAD.Close
Set objAD = Nothing
