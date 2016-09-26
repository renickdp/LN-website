<% if session("logonid")="" Then
Response.write "Palun logi sisse !"
Response.redirect ("default.asp?sid="&sid&"&action=login")
Else
'----------------------------------------------------------------------------------------------------------
t=request.querystring("t")
topage=Request.ServerVariables("HTTP_REFERER")
orn=request.querystring("orderno")
os=request.querystring("os")

if os="" Then os=0
sqlstmt = "SELECT orders.orderno,  Avg(orders.delivered) AS orderst, Avg(orders.accept) AS accept from orders  "
'sqlstmt = sqlstmt & " WHERE (((orders.username)='"&user&"'))"
sqlstmt = sqlstmt & " Group by orders.orderno "
if os=0 Then
os1=1
os2="Avatud tellimused"
sqlstmt = sqlstmt & " HAVING (((Avg(orders.delivered))<1))" 
Else
os1=0
os2="Suletud tellimused"
sqlstmt = sqlstmt & " HAVING (((Avg(orders.delivered))=1)) "
End If
sqlstmt = sqlstmt & " ORDER by orders.orderno desc "
  
rs.open sqlstmt, conn,1 ,1
%>
<table border="0" class="menutable" cellpadding="0" cellspacing="0"  width="180" style="border-collapse: collapse" bordercolor="#111111">
  <tr>
    <th width="180">
     <font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp; 
    <img src="img/arrow.gif" width="10" height="10">&nbsp;</font><a href="adm.asp?sid=<%=sid%>&orderno=<%=orn%>&os=<%=os1%>"><%=os2%></a>
     <font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;<img src="img/arrow.gif" width="10" height="10">&nbsp;</font></th>
  </tr>
  <%   
If rs.eof then
response.write "<tr><td> Tellimuste loetelu puudub!</td></tr>"
'response.write "<tr><th>&nbsp;</th></tr>"

'response.end
Else
  
If Not rs.EOF Then 
 	rs.MoveFirst
 %>
<tr>
        <td width="180">&nbsp; 
    <img src="img/arrow.gif" width="10" height="10">&nbsp;<a href="adm.asp?sid=<%=sid%>&os=<%=os%>">Tellimuste loetelu</a></td>
              </tr>

     
      <%
trow="#FCFBFF"
 Do while not rs.eof 
orderno=rs("orderno")
orderst=rs("orderst")
 accept=rs("accept")
if accept = 1 Then
accept="<b> ok</b> "
Else
accept="<b> ?</b> "
End if

 %>
   <tr>
        <td width="180">&nbsp; 
    <img src="img/arrow.gif" width="10" height="10">&nbsp;<a href="adm.asp?sid=<%=sid%>&os=<%=os%>&orderno=<%=orderno%>"><%=orderno%></a><%=accept%><% if orderst>0 and orderst <1 Then Response.write " *"%></td>
              </tr>
      <%
rs.MoveNext
if trow="#FCFBFF" Then	   	   
trow="#FFFFFF"
Else
trow="#FCFBFF"
End if
loop 
%>
    
   
<%End If
End If%>
   <tr>
        <th width="100%" bgcolor="#E6E6E6">&nbsp;</th>
     </tr>
   </table>
<p>&nbsp;</p>

<%rs.Close
End If%>