<%
if session("admin")<>1 Then response.redirect("default.asp")
'----------------------------------------------------------------------------------------------------------
%>
<h2>Kasutajate haldus</h2>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="500" class="maintable">
  <tr>
    <th>Kontaktisiku nimi&nbsp;</th>
    <th>Firma, aadress&nbsp;</th>
    <th>Telefon&nbsp;</th>
    <th>AH %&nbsp;</th>
  </tr>
<%
SQL = "SELECT* FROM users "
rs.Open SQL, Conn

trow="#FCFBFF"
Do while not rs.eof 
id=rs("ID")
pass=rs("pass")
login=rs("login")
fullname=rs("fullname")
coaddress = rs("coaddress")
coname=rs("coname")
phone=rs("phone")
discount=rs("discount")
description=rs("fullname")
%>
<tr>
    <td bgcolor="<%=trow%>"><a href="mailto:<%=login%>"><%=fullname%></a>&nbsp;</td>
    <td bgcolor="<%=trow%>"><%=coname%>&nbsp;</td>
    <td bgcolor="<%=trow%>"><%=phone%>&nbsp;</td>
    <td bgcolor="<%=trow%>" align="center"><b><a href="reg.asp?action=discount&id=<%=id%>&d=<%=discount-1%>">-</a>&nbsp;<%=discount%>&nbsp;<a href="reg.asp?action=discount&id=<%=id%>&d=<%=discount+1%>">+</a></b></td>
  </tr>
<tr>
    <td align="center" bgcolor="<%=trow%>"><font size="1">|&nbsp;blokeeri konto &nbsp;|</font></td>
    <td bgcolor="<%=trow%>" colspan="3"><font size="1"><%=coaddress%></font></td>
  </tr>

<%
if trow="#FCFBFF" Then	   	   
trow="#FFFFFF"
Else
trow="#FCFBFF"
End if

rs.MoveNext
 loop
 rs.close
 
'-----------------------------------------------------------------------------------
%>

</table>
<p>&nbsp;</p>