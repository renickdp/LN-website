<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0"  width="770" style="border-collapse: collapse" bordercolor="#111111" class="mainmenutable">
  <tr>
    <td valign="top" colspan="2" bgcolor="#FFFFCC">
    &nbsp;</td>  
    <td valign="top" colspan="2" align="right" bgcolor="#FFFFCC">
    <%if Session("logonid") = ""  Then %>&nbsp;<a href="<%=mainpage%>?sid=<%=sid%>&action=register&m=3">Register</a>&nbsp;<%End if%><a href="<%=mainpage%>?sid=<%=sid%>&action=help&m=3">Help</a>&nbsp;&nbsp;<%if Session("logonid") = ""  Then %><a href="<%=mainpage%>?sid=<%=sid%>&action=login&m=3">Log 
    in</a>&nbsp;&nbsp;<%Else %><%="&nbsp;"&fullname &  "&nbsp;:&nbsp;<a href='logout.asp'> Log off"%></a>&nbsp;<%End If%><%if session("admin")=1 Then %>&nbsp;<a href="adm.asp">Orders</a>&nbsp;&nbsp;<a href="adm_prod.asp">Products</a>&nbsp;&nbsp;<a href="adm_usr.asp">Users</a>&nbsp;<%End If%><%if discount >0 Then response.write "<font color='#ff9933' size='1'><b> - "& discount&" %</b></font>"%></td>  
 </tr>
  <tr>
    <td height="2" valign="top" colspan="3" bgcolor="#FFFFCC"></td>
    <td height="2" valign="top" bgcolor="#FFFFCC"></td>
  </tr>
  <tr>
    <td valign="top" colspan="4" bgcolor="#FFFFCC">
    <p align="left">
     <a href="<%=homepage%>?sid=<%=sid%>&">&nbsp;Home&nbsp;</a><a href="<%=mainpage%>?sid=<%=sid%>&m=1&group=<%=server.URLencode(request.querystring("group"))%>">&nbsp;Products&nbsp;</a><% if qty>0 Then %><a href="<%=mainpage%>?sid=<%=sid%>&action=viewcart">&nbsp;Cart&nbsp;</a><%response.write "<font size='1'>&nbsp;:&nbsp;" &session("qty")&"&nbsp;tk</font>"%><%End If%><%If session("logonid")<>"" Then%><a href="order.asp?sid=<%=sid%>&">
    &nbsp;Your orders&nbsp;</a><a href="<%=mainpage%>?sid=<%=sid%>&action=account&m=3">&nbsp;Your 
     Data&nbsp;</a><%End If%><font size="1">&nbsp;<%=date()%>&nbsp;&nbsp;</font></td>
  </tr>
  <tr>
    <td height="2" bgcolor="#FF9900" valign="top" colspan="3"></td>
    <td height="2" bgcolor="#FF9900" valign="top"></td>
  </tr>
<form method="POST" name="searchform" action="<%=mainpage&"?sid="&sid%>&amp;m=<%=m%>">
  <tr>
    <td nowrap bgcolor="#add8e6"><font color="#CC3300"><a name="0"></a>Search for&nbsp;:&nbsp;<input type="text" name="src" size="15" class="boxrequired" value="<%=request.form("src")%>">&nbsp;<input type="image" src="img/btn.gif" value="Submit" name="I1" border="0" align="bottom">&nbsp;</font></td>
    <td align="right" colspan="3" nowrap bgcolor="#add8e6">
   <%if m=1 Then %><font color="#CC3300"> </font><font color="#CC3300">
   <a href="<%=mainpage%>?sid=<%=sid%>&m=2&group=<%=server.URLencode(request.querystring("group"))%>">
   <font color="#CC3300">Product&nbsp;</font></a>:&nbsp;<select size="1" name="group" class="boxrequired" onchange="searchform.submit()">
     <option selected>---your choice---</option>
<% sqlstmt = "SELECT items.group from items  "
sqlstmt=sqlstmt & " WHERE len(group)>1 "
sqlstmt = sqlstmt & "GROUP by items.Group "
sqlstmt = sqlstmt & "ORDER by items.Group "
rs.open sqlstmt, conn,1 ,1  
If Not rs.EOF Then
rs.MoveFirst
while not rs.eof 
group = rs("group")
if len(group)>15 Then grp1=Left(group,15)&"..."
 %>   
     <option value="<%=group%>"><%=group%></option>
<%rs.MoveNext
wend 
End If
rs.Close
%>
 </select>
 <%Else%>
   <a href="<%=mainpage%>?sid=<%=sid%>&m=1&group=<%=server.URLencode(request.querystring("group"))%>">
   <font color="#CC3300">Brand&nbsp;:</font></a>&nbsp;<select size="1" name="product" class="boxrequired" onchange="searchform.submit()">
     <option selected>---your choice---</option>
<% sqlstmt = "SELECT items.product from items  "
sqlstmt=sqlstmt & " WHERE len(product)>1 "
sqlstmt = sqlstmt & "GROUP by items.product "
sqlstmt = sqlstmt & "ORDER by items.product "
rs.open sqlstmt, conn,1 ,1  
If Not rs.EOF Then
rs.MoveFirst
while not rs.eof 
product = rs("product")
if len(product)>15 Then prd1=Left(product,15)&"..."
 %>   
     <option value="<%=product%>"><%=product%></option>
<%rs.MoveNext
wend 
End If
rs.Close
%>
 </select>
 <% End if%>&nbsp;</font></td>
  </tr>
  <input type="hidden" name="m" value="<%=m%>">
</form>
<tr>
    <td height="2" bgcolor="#FF9900" valign="top" colspan="3"></td>
    <td height="2" bgcolor="#FF9900"  valign="top"></td>
  </tr>
</table></center>
</div>