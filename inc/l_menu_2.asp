<%
m=request.querystring("m")
group=request.querystring("group")
group1=request.querystring("group1")
group2=request.querystring("group2")
group3=request.querystring("group3")
src=request.querystring("src")
if m=2 or m="" Then m=1
if m=1 Then
%>
<table width="180" class="menutable" cellspacing="0" cellpadding="0" border="0">
<tr>
        <th width="100%" bgcolor="#E6E6E6"><img border="0" src="img/arrow.gif" alt=" " width="10" height="11">&nbsp;Products&nbsp;<img border="0" src="img/arrow.gif" alt=" " width="10" height="11"></th>
     </tr>
<% '-LEVEL 1 ----------------------------------------------------------------------------------------------
if group="" Then 
sqlstmt = "SELECT items.GROUP FROM items GROUP BY items.GROUP "
sqlstmt = sqlstmt & "ORDER by items.Group "
rs.open sqlstmt, conn,1 ,1  
while not rs.eof 
group = rs("group")
if group<>"" or group<>" " Then 
gr=group
if len(gr)>15 Then gr=left(gr,15)&"..."
%> 
<tr>
<td width="100%" onmouseover="this.style.backgroundColor='#FFE7CE';" onmouseout="this.style.backgroundColor='';">&nbsp;&nbsp;<img border="0" src="img/arrow.gif" alt=" " width="10" height="11">&nbsp;<a title="<%=group%>" href="<%=p_url%>?sid=<%=sid%>&m=<%=m%>&group=<%=server.URLencode(group)%>"><%=gr%></a></td>     
</tr>	
<% End if
rs.MoveNext
   wend 
   rs.Close
Else
gr=group
if len(gr)>15 Then gr=left(gr,15)&"..."
%>
<tr>
<td width="100%" onmouseover="this.style.backgroundColor='#FFE7CE';" onmouseout="this.style.backgroundColor='';">&nbsp;<img border="0" src="img/arrow.gif" alt=" " width="10" height="11">&nbsp;<a href="<%=p_url%>?sid=<%=sid%>&m=<%=m%>">Show 
all</a></td>
</tr>	
<tr>
<td width="100%" onmouseover="this.style.backgroundColor='#FFE7CE';" onmouseout="this.style.backgroundColor='';">&nbsp;&nbsp;<img border="0" src="img/arrow.gif" alt=" " width="10" height="11">&nbsp;<a title="<%=group%>" href="<%=p_url%>?sid=<%=sid%>&m=<%=m%>&group=<%=server.URLencode(group)%>"><%=gr%></a></td>
</tr>

<%  '-LEVEL 2 ----------------------------------------------------------------------------------------------
group1=request.querystring("group1")

  if group1="" or group1=" "  Then
sqlstmt = "SELECT items.GROUP1 FROM items "
sqlstmt = sqlstmt & " WHERE items.GROUP='"& group & "' "
sqlstmt = sqlstmt & " GROUP BY items.GROUP1 "
sqlstmt = sqlstmt & "ORDER by items.Group1 "

rs.open sqlstmt, conn,1 ,1  
while not rs.eof 
gp1 = rs("group1")
if gp1<>"" or gp1<>" " Then 
gr1=gp1
if len(gr1)>15 Then gr1=left(gr1,15)&"..."
%>
<tr>
<td width="100%" onmouseover="this.style.backgroundColor='#FFE7CE';" onmouseout="this.style.backgroundColor='';">&nbsp;&nbsp;&nbsp;&nbsp;<img border="0" src="img/arrow.gif" alt=" " width="10" height="11">&nbsp;<a title="<%=group1%>" href="<%=p_url%>?sid=<%=sid%>&m=<%=m%>&group=<%=server.URLencode(group)%>&group1=<%=server.URLencode(gp1)%>"><%=gr1%></a></td>
</tr>	
<%  End if
	rs.MoveNext
   	wend 
    rs.Close
Else
gr1=group1
if len(gr1)>15 Then gr1=left(gr1,15)&"..."


%>
<tr>
<td width="100%" onmouseover="this.style.backgroundColor='#FFE7CE';" onmouseout="this.style.backgroundColor='';">&nbsp;&nbsp;&nbsp;&nbsp;<img border="0" src="img/arrow.gif" alt=" " width="10" height="11">&nbsp;<a title="<%=group1%>" href="<%=p_url%>?sid=<%=sid%>&m=<%=m%>&group=<%=server.URLencode(group)%>&group1=<%=server.URLencode(group1)%>"><%=gr1%></a></td>
</tr>

<%End if
End If 
 '-LEVEL 3 ----------------------------------------------------------------------------------------------
group2=request.querystring("group2")
if group2="" or group2=" " Then
sqlstmt = "SELECT items.GROUP2 FROM items "
sqlstmt = sqlstmt & " WHERE items.GROUP='"& group & "' AND items.GROUP1='"& group1 & "' "
sqlstmt = sqlstmt & " GROUP BY items.GROUP2 "
sqlstmt = sqlstmt & "ORDER by items.Group2 "

rs.open sqlstmt, conn,1 ,1  
while not rs.eof 
gp2 = rs("group2")
if gp2<>"" or gp2<>" " Then 
gr2=gp2
if len(gr2)>15 Then gr2=left(gr2,15)&"..."
%>
<tr><td width="100%" onmouseover="this.style.backgroundColor='#FFE7CE';" onmouseout="this.style.backgroundColor='';">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img border="0" src="img/arrow.gif" alt=" " width="10" height="11">&nbsp;<a title="<%=group2%>" href="<%=p_url%>?sid=<%=sid%>&m=<%=m%>&group=<%=server.URLencode(group)%>&group2=<%=server.URLencode(gp2)%>&group1=<%=server.URLencode(group1)%>"><%=gr2%></a></td>
</tr>	
<% End if
rs.MoveNext
   wend 
   rs.Close
Else
gr2=group2
if len(gr2)>15 Then gr2=left(gr2,15)&"..."
%>
<tr>
<td width="100%" onmouseover="this.style.backgroundColor='#FFE7CE';" onmouseout="this.style.backgroundColor='';">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img border="0" src="img/arrow.gif" alt=" " width="10" height="11">&nbsp;<a title="<%=group2%>" href="<%=p_url%>?sid=<%=sid%>&m=<%=m%>&group=<%=server.URLencode(group)%>&group1=<%=server.URLencode(group1)%>&group2=<%=server.URLencode(group2)%>"><%=gr2%></a></td>
 </tr>
<%
End if
'End If 

if group3=""  Then
sqlstmt = "SELECT items.GROUP3 FROM items "
sqlstmt = sqlstmt & " WHERE items.GROUP='"& group & "' AND items.GROUP1='"& group1 & "' AND items.GROUP2='"& group2 & "' "
sqlstmt = sqlstmt & " GROUP BY items.GROUP3 "
sqlstmt = sqlstmt & "ORDER by items.Group3 "

rs.open sqlstmt, conn,1 ,1  
while not rs.eof 
gp3 = rs("group3")
if gp3<>"" or gp3<>" " Then 
gr3=gp3
if len(gr3)>15 Then gr3=left(gr3,15)&"..."
%>
<tr><td width="100%" onmouseover="this.style.backgroundColor='#FFE7CE';" onmouseout="this.style.backgroundColor='';">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img border="0" src="img/arrow.gif" alt=" " width="10" height="11">&nbsp;<a title="<%=gp3%>" href="<%=p_url%>?sid=<%=sid%>&m=<%=m%>&group=<%=server.URLencode(group)%>&group2=<%=server.URLencode(group2)%>&group1=<%=server.URLencode(group1)%>&group3=<%=server.URLencode(gp3)%>"><%=gr3%></a></td>
</tr>	
<% End if
rs.MoveNext
   wend 
   rs.Close
Else
gr3=group3
if len(gr3)>15 Then gr3=left(gr3,15)&"..."
%>
<tr>
<td width="100%" onmouseover="this.style.backgroundColor='#FFE7CE';" onmouseout="this.style.backgroundColor='';">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img border="0" src="img/arrow.gif" alt=" " width="10" height="11">&nbsp;<a title="<%=group3%>" href="<%=p_url%>?sid=<%=sid%>&m=<%=m%>&group=<%=server.URLencode(group)%>&group1=<%=server.URLencode(group1)%>&group2=<%=server.URLencode(group2)%>&group3=<%=server.URLencode(group3)%>"><%=gr3%></a></td>
</tr>
<%End if%>
<tr><td width="100%">&nbsp;</td></tr>
</table>

<%End if%> 
<%if m=3 Then %>
<table width="180" class="menutable" cellspacing="0" cellpadding="0" border="0">
<tr>
        <th width="100%" bgcolor="#E6E6E6"><img border="0" src="img/arrow.gif" alt=" " width="10" height="11">&nbsp;How 
        to use&nbsp;<img border="0" src="img/arrow.gif" alt=" " width="10" height="11"></th>
     </tr>
 <tr>
        <td width="100%" onmouseover="this.style.backgroundColor='#FFE7CE';" onmouseout="this.style.backgroundColor='';" >&nbsp;<img border="0" src="img/arrow.gif" alt=" " width="10" height="11">
        <a href="default.asp?sid=<%=sid%>&action=help&m=3#1">Register</a></a></td>
     </tr>
<tr>
        <td width="100%" onmouseover="this.style.backgroundColor='#FFE7CE';" onmouseout="this.style.backgroundColor='';" >&nbsp;<img border="0" src="img/arrow.gif" alt=" " width="10" height="11">
        <a href="default.asp?sid=<%=sid%>&action=help&m=3#5">Products</a></a></td>
     </tr>
<tr>
        <td width="100%" onmouseover="this.style.backgroundColor='#FFE7CE';" onmouseout="this.style.backgroundColor='';" >&nbsp;<img border="0" src="img/arrow.gif" alt=" " width="10" height="11">
        <a href="<%=mainpage%>?sid=<%=sid%>&action=help&m=3#6">Cart</a></td>
     </tr>
<tr>
        <td width="100%" onmouseover="this.style.backgroundColor='#FFE7CE';" onmouseout="this.style.backgroundColor='';" >&nbsp;<img border="0" src="img/arrow.gif" alt=" " width="10" height="11">
        <a href="<%=mainpage%>?sid=<%=sid%>&action=help&m=3#7">Make your order</a></a></td>
     </tr>
<tr>
        <td width="100%" onmouseover="this.style.backgroundColor='#FFE7CE';" onmouseout="this.style.backgroundColor='';" >&nbsp;<img border="0" src="img/arrow.gif" alt=" " width="10" height="11">
        <a href="<%=mainpage%>?sid=<%=sid%>&action=help&m=3#10">Order status</a></td>
     </tr>
<tr>
        <td width="100%" onmouseover="this.style.backgroundColor='#FFE7CE';" onmouseout="this.style.backgroundColor='';" >&nbsp;</td>
     </tr>
<tr>
        <td width="100%" onmouseover="this.style.backgroundColor='#FFE7CE';" onmouseout="this.style.backgroundColor='';" >&nbsp;</td>
     </tr>
<tr>
        <td width="100%" onmouseover="this.style.backgroundColor='#FFE7CE';" onmouseout="this.style.backgroundColor='';" >&nbsp;</td>
     </tr>
<tr>
        <td width="100%" onmouseover="this.style.backgroundColor='#FFE7CE';" onmouseout="this.style.backgroundColor='';" >&nbsp;</td>
     </tr>
  <tr>
        <td width="100%">&nbsp;</td>
     </tr>

</table>
<%End if%>