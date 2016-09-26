<!--#INCLUDE FILE="inc/general.asp"-->
<%

orderno=request.querystring("orderno")
code=request.querystring("code")

sAction=Request.QueryString("action")

Select Case sAction
case "accept"
y=request.querystring("y")
sqlstmt = "UPDATE orders"
  sqlstmt = sqlstmt & " SET "
  sqlstmt = sqlstmt & "accept='" & y & "'"
  sqlstmt = sqlstmt & " WHERE orderno=" & orderno  

Response.write SQLstmt &"<br>"
Set RS = conn.execute(SQLstmt)

Response.redirect ("adm.asp?")

case "deliver"

x=request.querystring("x")
sqlstmt = "UPDATE orders"
  sqlstmt = sqlstmt & " SET "
  sqlstmt = sqlstmt & "delivered='" & x & "'"
  if code<>"" Then
  sqlstmt = sqlstmt & " WHERE orderno=" & orderno & " and code='" & code &"' " 
  Else
  sqlstmt = sqlstmt & " WHERE orderno=" & orderno 
  End if

Response.write SQLstmt &"<br>"
Set RS = conn.execute(SQLstmt)
Response.redirect ("adm.asp?")

End Select
%>