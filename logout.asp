<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1257">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Thanx</title>
<%
Response.Cookies("passes").Expires = date()-1
Response.Cookies("passes2").Expires = date()-1
response.cookies("login").Expires = date()-1

Session("logonid") = ""
Session("admin") = ""
session("discount")=0

Response.Redirect("default.asp")
'Session.abandon
%>
</head>

<BODY onload="setTimeout('top.close()',3000)">

<p align="center">
<br>
&nbsp;</p>

</body>

</html>