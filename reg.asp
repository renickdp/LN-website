<!--#INCLUDE FILE="inc/general.asp"-->
<%
sAction=Request.QueryString("action")

id = request.form("id")
login = request.form("login")
if sAction <> "pwd" and sAction <> "discount" Then

if sAction="register" Then
SQL = "SELECT login FROM users Where login='"& login &"' "
rs.Open SQL, Conn
If rs.eof then
Else
Response.write login&" selline e posti aadress juba eksisteerib andmebaasis"
response.write "<form>"
	response.write "<input type='button' value='Retry' onclick=history.back()>"
	response.write "</form>"
Response.end
End If
rs.close

If instr(login, "@") = 0 OR instr(login, ".") = 0 OR Len(login) < 7 Then
	response.write "<center>Palun sisesta korrektne e-maili aadress!"
	response.write "<form>"
	response.write "<input type='button' value='Retry' onclick=history.back()>"
	response.write "</form>"
	response.end
End If
End If 

fullname = request.form("fullname")
If IsEmpty(request.form("fullname"))or len(fullname)<3 or request.form("fullname")="" then
	response.write "<center>Palun sisesta nimi."
	response.write "<form>"
	response.write "<input type='button' value='Retry' onclick=history.back()>"
	response.write "</form>"
	response.end
Else
fullname = request.form("fullname")
End If

coname = request.form("coname")
If IsEmpty(request.form("coname")) or len(coname)<3 or request.form("coname")="" then
	response.write "<center>Palun sisesta nimi."
	response.write "<form>"
	response.write "<input type='button' value='Retry' onclick=history.back()>"
	response.write "</form>"
	response.end
Else
coname = request.form("coname")
End If

coaddress = request.form("coaddress")
If IsEmpty(request.form("coaddress"))  or len(coaddress)<5 or request.form("coaddress")="" then
	response.write "<center>Palun sisesta aadress."
	response.write "<form>"
	response.write "<input type='button' value='Retry' onclick=history.back()>"
	response.write "</form>"
	response.end
Else
fullname = request.form("fullname")
End If

phone = request.form("phone")
If IsEmpty(request.form("phone"))  or len(phone)<6 or request.form("phone")="" then
	response.write "<center>Palun sisesta telefon."
	response.write "<form>"
	response.write "<input type='button' value='Retry' onclick=history.back()>"
	response.write "</form>"
	response.end
Else
phone = request.form("phone")
End If

if sAction="register" Then 
 
pass1 = request.form("pass1")
If IsEmpty(request.form("pass1")) or request.form("pass1")="" then
	response.write "<center>Palun sisesta salasõna."
	response.write "<form>"
	response.write "<input type='button' value='Retry' onclick=history.back()>"
	response.write "</form>"
	response.end
Else
pass1 = request.form("pass1")
End If

pass2 = request.form("pass2")
If IsEmpty(request.form("pass2")) or request.form("pass2")="" or pass2<>pass1 then
	response.write "<center>Palun sisesta salasõna uuesti."
	response.write "<form>"
	response.write "<input type='button' value='Retry' onclick=history.back()>"
	response.write "</form>"
	response.end
Else
pass = pass2
End If

End If
End if

Select Case sAction
Case "register"

SQLstmt = "INSERT INTO users (login,pass,fullname,coname,coaddress,phone,adate)"
SQLstmt = SQLstmt & " VALUES ("
SQLstmt = SQLstmt & "'" & login & "',"
SQLstmt = SQLstmt & "'" & pass & "',"
SQLstmt = SQLstmt & "'" & fullname & "',"
SQLstmt = SQLstmt & "'" & coname & "'," 
SQLstmt = SQLstmt & "'" & coaddress & "',"
SQLstmt = SQLstmt & "'" & phone & "'," 
SQLstmt = SQLstmt & "'" & date() & "'"
SQLstmt = SQLstmt & ")"
	
Response.write SQLstmt &"<br>"
Set RS = conn.execute(SQLstmt)

Response.redirect ("default.asp?action=login&login="&login)

Case "change"

sqlstmt = "UPDATE USERS"
  sqlstmt = sqlstmt & " SET "
  sqlstmt = sqlstmt & " fullname='" & fullname & "'," 
  sqlstmt = sqlstmt & " coname='" & coname & "',"
  sqlstmt = sqlstmt & "coaddress='" & coaddress & "',"
  sqlstmt = sqlstmt & "phone='" & phone & "'"
  sqlstmt = sqlstmt & " WHERE id=" & id 


Response.write SQLstmt &"<br>"
Set RS = conn.execute(SQLstmt)
Response.redirect ("default.asp?action=account")

case "pwd"

pass0 = request.form("pass0")
pass = request.form("pass")

If IsEmpty(request.form("pass0")) or request.form("pass0")="" or pass0 <> pass Then 
response.write "<center>Salasõna ?"
	response.write "<form>"
	response.write "<input type='button' value='Retry' onclick=history.back()>"
	response.write "</form>"
	response.end
End If

pass1 = request.form("pass1")
If IsEmpty(request.form("pass1")) or request.form("pass1")="" then
	response.write "<center>Palun sisesta salasõna."
	response.write "<form>"
	response.write "<input type='button' value='Retry' onclick=history.back()>"
	response.write "</form>"
	response.end
Else
pass1 = request.form("pass1")
End If

pass2 = request.form("pass2")
If IsEmpty(request.form("pass2")) or request.form("pass2")="" or pass2<>pass1 then
	response.write "<center>Palun sisesta salasõna uuesti."
	response.write "<form>"
	response.write "<input type='button' value='Retry' onclick=history.back()>"
	response.write "</form>"
	response.end
Else
pass = pass2
End If

sqlstmt = "UPDATE USERS"
  sqlstmt = sqlstmt & " SET "
  sqlstmt = sqlstmt & "pass='" & pass & "'"
  sqlstmt = sqlstmt & " WHERE id=" & id 

Response.write SQLstmt &"<br>"
Set RS = conn.execute(SQLstmt)

' insert e mail 


Response.redirect ("default.asp?action=account")

case "discount"
id=request.querystring("id")
discount=request.querystring("d")
if discount<=0 then discount = 0
if discount >= 15 Then discount=15
sqlstmt = "UPDATE USERS"
  sqlstmt = sqlstmt & " SET "
  sqlstmt = sqlstmt & "discount='" & discount & "'"
  sqlstmt = sqlstmt & " WHERE id=" & id 


Response.write SQLstmt &"<br>"
Set RS = conn.execute(SQLstmt)
Response.redirect ("adm_usr.asp")

End Select
%>