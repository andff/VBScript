' Replace with your sender's email address
strSenderEmail = "seu_email@seudominio.com.br"

' Replace with your sender's password
strSenderPassword = "sua_senha"

' Replace with the recipient's email address
strRecipientEmail = "destinatario@seudominio.com.br"

' Set the email subject
strSubject = "Assunto do seu email"

' Set the email body
strBody = "Corpo do seu email. Você pode incluir HTML aqui também."

' Create a CDO.Message object
Set objEmail = CreateObject("CDO.Message")

' Configure the email properties
objEmail.From = strSenderEmail
objEmail.To = strRecipientEmail
objEmail.Subject = strSubject
objEmail.TextBody = strBody

' Configure the SMTP server settings
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' Use SMTP
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.seudominio.com.br" ' Replace with your SMTP server address
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 ' Replace with your SMTP server port (usually 25 or 587)

' If your SMTP server requires authentication, uncomment the following lines and replace with your credentials
' objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthentication") = 1 ' Enable authentication
' objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/username") = strSenderEmail
' objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/password") = strSenderPassword

' Send the email
objEmail.Send

' Display a success message
MsgBox "Email enviado com sucesso!"
