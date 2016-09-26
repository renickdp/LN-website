<%
sqlstmt = "SELECT specials.code, specials.description, items.PRICE FROM items INNER JOIN specials ON items.CODE = specials.code "
sqlstmt = sqlstmt &  " WHERE (specials.startdate <= datevalue(date()) AND specials.enddate >= datevalue(date()) ) "
rs.open sqlstmt, conn,1 ,1
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
  <tr><td width="100%">
<ul>
<%
If rs.eof then
response.write "<li><a href='" & mainpage &"'>Head hinnad ... </a></li>"

else

Do while not rs.eof 
code = rs("code")
price=rs("price")
price=price-((price*discount)/100)
price=price*vat
description=rs("description")
Response.write "<li><font size='1'><a href='" & mainpage &"?code="& server.urlencode(code) & "'>"&description & "</a><b></font><font color='cc3300' size='1'> " & formatnumber(price,2) & "</b></font></li>"

rs.MoveNext
loop
rs.close
End If
%>
</ul>
</td>
    
  </tr>
</table>