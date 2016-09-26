<!--#INCLUDE FILE="inc/general.asp"-->
<% 
Response.Buffer = True
sid=(Session.SessionID)    
l=Session("login")
id=Request.form("id")
if len(id)>0 then id=replace(id,chr(39),"''") 
pwd=Request.form("pwd")
if len(pwd)>0 then pwd=replace(pwd,chr(39),"''") 

if l = "" Then l=0

function ValidateLogin(sId,sPwd) 
Dim Apples
 
SQLtemp = "SELECT * FROM users WHERE login = '" & id & "' "
RS.Open SQLtemp, conn, 1,1
while not rs.eof

dim user
user = rs("login") ' emailaddress
dim fullname
fullname = rs("fullname")

response.cookies("passes") = user
response.cookies("passes2") = fullname

ValidateLogin = False
If id = rs("login") AND pwd = rs("pass") Then
Session("admin")=rs("id")
Session("discount")=rs("discount")
ValidateLogin = True 
Else
ValidateLogin = False
End If

rs.MoveNext
Wend
rs.Close

set Apples = Nothing

End function
 
Dim sText
  Dim fBack
    fBack = False
    if Request.Form("dologin") = "yes" Then 
    	'Try To login
    	if ValidateLogin( id, pwd) = True Then
    		'It is OK!!!
    		fBack = True
   		Session("logonid") = id
   		Session("login") = 0
    	Response.Redirect(Request.ServerVariables("HTTP_REFERER"))		
    	Else
    	sText = "Wrong ID or Password"
 		Session("login") = l+1
      	End if
    Else
  	'We are Not trying To login...
    	if Session("logonid") <> "" Then 
      		fBack = True
    		'We are logged In so lets go back To the file that included us 
	    	 	Else
    		sText = "Log In !"
    	End if
    End if
    if fBack = False Then 
    %>
    <%
    if Request.form ("siteid")<>"" Then
    Response.Redirect(Request.form ("siteid"))
    Session("logonid") = ""
	Response.end
	End If
	end if
	%>