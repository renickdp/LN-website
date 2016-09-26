<!--#INCLUDE FILE="inc/general.asp"-->
<%

code=request.querystring("code")

sAction=Request.QueryString("action")

Select Case sAction
case "yee"
y=request.querystring("y")
sqlstmt = "UPDATE items"
  sqlstmt = sqlstmt & " SET "
  sqlstmt = sqlstmt & "yee='" & y & "'"
  sqlstmt = sqlstmt & " WHERE code='" & code &"' " 

Response.write SQLstmt &"<br>"
Set RS = conn.execute(SQLstmt)
Conn.Close
Response.redirect ("adm_prod.asp?action=products&code="&code)

case "discount"

case "addspecial"

description=request.form("description")
if description ="" Then response.redirect(request.servervariables("HTTP_REFFERER"))
code=request.form("code")
if code ="" Then response.redirect(request.servervariables("HTTP_REFFERER"))
startdate=request.form("startdate")
if startdate ="" Then response.redirect(request.servervariables("HTTP_REFFERER"))
enddate=request.form("enddate")
if enddate ="" Then response.redirect(request.servervariables("HTTP_REFFERER"))

SQLstmt = "INSERT INTO specials (description,code,startdate,enddate)"
SQLstmt = SQLstmt & " VALUES ("
SQLstmt = SQLstmt & "'" & description & "',"
SQLstmt = SQLstmt & "'" & code & "',"
SQLstmt = SQLstmt & "'" & startdate & "'," 
SQLstmt = SQLstmt & "'" & enddate & "'"
SQLstmt = SQLstmt & ")"

Response.write SQLstmt &"<br>"
Set RS = conn.execute(SQLstmt)
Conn.Close
Response.redirect ("adm_prod.asp?action=specials")

case "delspecial"

if code ="" Then response.redirect(request.servervariables("HTTP_REFFERER"))
SQLstmt = "DELETE from specials where code='" &code & "'" 
Set RS = conn.execute(SQLstmt)
Conn.Close
Response.redirect ("adm_prod.asp?action=specials")


End Select
%>