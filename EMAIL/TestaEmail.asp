<html>
<head>
</head>
<body>

<br> start2


<%


dim smtpserver, youremail, yourpassword
smtpserver = "smtp.mailtrap.io"
youremail = "238c1b3157f367"
yourpassword = "cc3a93c8c22708" 

Dim ObjSendMail
Set ObjSendMail = CreateObject("CDO.Configuration")
Set objSendmsg =  CreateObject ("CDO.Message") 
ObjSendMail.Fields ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
ObjSendMail.Fields ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpserver
ObjSendMail.Fields ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 2525
'ObjSendMail.Fields ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true 'Use SSL for the connection
ObjSendMail.Fields ("http://schemas.microsoft.com/cdo/configuration/sendtls") = true
ObjSendMail.Fields ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
ObjSendMail.Fields ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
ObjSendMail.Fields ("http://schemas.microsoft.com/cdo/configuration/sendusername") = youremail
ObjSendMail.Fields ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = yourpassword
ObjSendMail.Fields.Update
set objSendmsg.Configuration = ObjSendMail



objSendmsg.To = "suporte.tedesco@tr.com"
'ObjSendMail.CC = "suporte.tedesco@tr.com"
objSendmsg.Subject = "TESTE EMAIL"
objSendmsg.From = "sustentacaotr@gmail.com"
objSendmsg.TextBody = ""
objSendmsg.HTMLBody = "<b>ENVIADO - " & now() & "</b>"
objSendmsg.Send
response.Write "Mail Sent Successfully!<br />Check your Inbox for: "
Set ObjSendMail = Nothing 
set objSendmsg = nothing 
%>



</body>

</html>
