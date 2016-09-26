<!--#INCLUDE FILE="inc/general.asp"-->

<%

login=request.form("remind")
If instr(login, "@") = 0 OR instr(login, ".") = 0 OR Len(login) < 7 Then
	response.write "<center>Please type correct email address!"
	response.write "<form>"
	response.write "<input type='button' value='Retry' onclick=history.back()>"
	response.write "</form>"
	response.end
End If

SQL = "SELECT login, pass FROM users Where login='"& login &"' "
rs.Open SQL, Conn

If Not rs.EOF Then
'rs.MoveFirst
Do while not rs.eof
pass=rs("pass")
rs.Movenext
loop

mailserver="mail.mail.com:25"

dim mailbody
saatja="info@ln.com"
username="info@ln.com"

mailbody = " LM Shop userid and password " & Vbcrlf & vbcrlf
mailbody = mailbody & " userid : " & login & Vbcrlf  
mailbody = mailbody & " password : " & pass & Vbcrlf & vbcrlf 
mailbody = mailbody & vbCrlf & vbcrlf &  "http://ln.com" 
blabla = "reminder "
Response.Write mailbody

set SMTP=Server.CreateObject("Jmail.SMTPMail")
SMTP.ServerAddress= mailserver
SMTP.Sender=saatja
SMTP.AddRecipient login
SMTP.AddRecipient username
SMTP.Subject=blabla
SMTP.Body=mailbody
SMTP.Execute
Session("login")=0
rs.close
Else 
rs.close
Response.write "<p align='center'>"&login&"<br> This e mail address does not exists in our database !</p>"
	'response.write "<form>"
	response.write "<p align='center'><a href='default.asp?action=register'>Register</a></p>"
	'response.write "<input type='button' value='Retry' onclick=history.back()>"
	'response.write "</form>"
Session("login")=0
Response.end

End If
Response.redirect("default.asp?action=login")

%>