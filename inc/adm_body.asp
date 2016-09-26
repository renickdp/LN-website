<h2>Order Management</h2>
<%
if session("admin")<>1 Then response.redirect("default.asp")
'----------------------------------------------------------------------------------------------------------
t=request.querystring("t")
pg=request.querystring("pg")
if pg="" Then pg=1

topage=Request.ServerVariables("HTTP_REFERER")
orn=request.querystring("orderno")
if orn<>"" Then 

sqlstmt = "SELECT * from orders  "
sqlstmt = sqlstmt & " WHERE orders.orderno = "& orn &" "
rs.open sqlstmt, conn,1 ,1
Do while not rs.eof 
Customer = rs("customer")
adate=rs("adate")
descr=rs("description")
shipping=rs("shipping")
rs.MoveNext
loop 
rs.Close


sqlstmt = "SELECT * from orders  "
sqlstmt = sqlstmt & " WHERE orders.orderno = "& orn &"  "
  
rs.open sqlstmt, conn,1 ,1
   
If rs.eof then
response.write "<center>Select a course from menu bar in the left"
response.write "<br></center>"
'response.end
Else
  
If Not rs.EOF Then 
 	rs.MoveFirst
 %>
<table border="0" class="maintable"  cellpadding="0" cellspacing="0"  width="500">
  <tr>
    <td>
    <table border="0" cellpadding="0" cellspacing="0" width="500">
      <tr>
        <td valign="top"><br><%Response.write "<H1>Tellimus : "&orn&"</H1>"%>
        <%=adate%>
        <br><b>
        <%=customer%></b>
        <br><br>
        <%if len(descr)>1 then response.write descr%>

        </td>
        <td valign="top">&nbsp;</td>
        <td valign="top">
        <p align="right"><img border="0" src="img/logo.gif" align="right"></td>
      </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td width="100%">
    <table border="0" class="maintable" cellpadding="0" cellspacing="0"  width="500">
      <tr>
        <td width="10" bgcolor="#F2F0FF" align="center"><b>&nbsp;?</b></td>
        <td width="300" bgcolor="#F2F0FF"><b>Kood</b>&nbsp;</td>
        <td width="20" bgcolor="#F2F0FF" align="center"><b>Kogus</b></td>
        <td width="50" bgcolor="#F2F0FF" align="right"><b>Hind</b></td>
        <td width="70" bgcolor="#F2F0FF" align="right"><b>Summa</b></td>
      </tr>
      <%
trow="#FCFBFF"
 Do while not rs.eof 
code = rs("code")
qty = rs("qty")
status=rs("delivered")
if status=0 Then txt="Töötluses"
if status=1 Then txt="Tarnitud"
price=rs("price")
comment=rs("comment")
 %>
   <tr>
        <td bgcolor="<%=trow%>" ><img border="0" alt="<%=txt%>" src="img/<%=status%>.gif"></td>
        <td bgcolor="<%=trow%>"><a href="default.asp?sid=<%=sid%>&code=<% =code%>&"><%=code%></a><font color="CCCCCC" title="<%=comment%>" size="1"><%=" " & left(comment,50)%></font></td>
        <td bgcolor="<%=trow%>" align="center"><%=qty%></td>
        <td bgcolor="<%=trow%>" align="right"><%=formatnumber(price,2)%></td>
        <td bgcolor="<%=trow%>" align="right"><b><%=formatnumber(qty*price,2)%></b></td>
      </tr>
      <% sTotal = sTotal + (qty * price)

rs.MoveNext
if trow="#FCFBFF" Then	   	   
trow="#FFFFFF"
Else
trow="#FCFBFF"
End if
loop 
rs.Close
If sTotal <> 0 Then
		sShipping = 0
		sShipping = shipping
	Else
		sShipping = 0
	End If
	sTotal = sTotal + sShipping
	sVAT = (sTotal) * 0.18
	stotalfin=stotal+svat

%>
    </table>
    </td>
  </tr>
  <tr>
    <td width="100%">
    <table border="0" cellpadding="0" class="maintable" cellspacing="0"  width="100%">
      <tr>
        <td>
        <p align="right"><b>&nbsp; </b></td>
        <td width="50%" align="right"><b><%=" Transport : " &formatnumber(sShipping,2)%></b></td>
      </tr>
      <tr>
        <td>
        <p align="right">&nbsp;</td>
        <td align="right"><b><%=" Kokku : " &formatnumber(sTotal,2)%></b></td>
      </tr>
      <tr>
        <td>
        <p align="right">&nbsp;</td>
        <td align="right"><b><%=" 18% : " &FormatNumber(sVAT,2) %></b></td>
      </tr>
      <tr>
        <td>
        <p align="right">&nbsp;</td>
        <td align="right"><b><%= "Kõik kokku : "& FormatNumber(stotalfin,2) %></b></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<%
End If
End If
Else 
%>
  <% 
sqlstmt = "SELECT orders.orderno, Last(orders.customer) AS customer, Last(orders.adate) AS adate, Sum((orders.qty)*(orders.price)) AS sumoford,Avg(orders.delivered) AS delivered ,Last(orders.username) AS username , Last(orders.description) as Description, Last(orders.shipcode) AS shipcode, Last(orders.accept) AS accept"
sqlstmt = sqlstmt & " FROM orders GROUP BY orders.orderno"
if os=0 Then
os1=1
os2="Active Orders"
sqlstmt = sqlstmt & " HAVING (((Avg(orders.delivered))<1))" 
Else
os1=0
os2="Order Management"
sqlstmt = sqlstmt & " HAVING (((Avg(orders.delivered))=1)) "
End If
%>
<table class="maintable" width="500" border="0" cellpadding="0" cellspacing="0">
<tr>
        <th bgcolor="#F2F0FF" colspan="6"><b><%=os2%>&nbsp;</b></th>             
 </tr>

<tr>
        <td bgcolor="#F2F0FF"><b>&nbsp;order</b></td><td bgcolor="#F2F0FF"><b></b></td>
<td bgcolor="#F2F0FF"><b>Date</b></td><td bgcolor="#F2F0FF"><b>Client</b></td><td bgcolor="#F2F0FF">
        <b>Total</b></td>
        <td bgcolor="#F2F0FF"><b>Info</b></td>             
 </tr>
<%
rs.open sqlstmt, conn,1 ,1
Const NrPerPage = 20
  If Request.QueryString("pg") = "" then
    CurrentPage = 1 'We're on the first page
Else
    CurrentPage = CInt(Request.QueryString("pg"))
End If
If Not rs.EOF Then 
rs.MoveFirst
 	rs.PageSize = NrPerPage
    TotalPages = rs.PageCount
    rs.AbsolutePage = CurrentPage
%>

<%
trow="#FCFBFF"
 Do while not rs.eof and Count1  < rs.PageSize
orderno=rs("orderno")
customer=rs("customer")
adate=rs("adate")
accept=rs("accept")
if accept=0 Then
y=1
y3="<b> ?</b> "
y2="aktsepteeri"
Else
y=0
y2="tühista"
y3="<b> ok</b> "
End if
delivered=rs("delivered")
if delivered<1 Then
x=1
x2="tarni kaup"
Else
x=0
x2="kaup tagasi"
End if

sum=(rs("sumoford"))
shipcode=rs("shipcode")
username=rs("username")
descr=rs("description")
if len(descr)>30 Then descr=Left(descr,30) & "..."
 %>
   <tr>
        <td bgcolor="<%=trow%>">&nbsp; <%=y3%>
    <img src="img/arrow.gif" width="10" height="10">&nbsp;<a href="adm.asp?sid=<%=sid%>&os=<%=os%>&orderno=<%=orderno%>"><font size="1"><%=orderno%></font></a></td>
<td bgcolor="<%=trow%>"><font size="1"><%=adate%></font></td><td bgcolor="<%=trow%>"><a href="mailto:<%=username%>"><font size="1"><%=customer%></font></a></td><td bgcolor="<%=trow%>"><font size="1"><b><%=Formatnumber(sum,2)%></b></font></td><td bgcolor="<%=trow%>"><font size="1"><%=descr%></font></td>
        <td bgcolor="#FFE7CE"><font size="1"><a href="ord.asp?action=accept&orderno=<%=orderno%>&y=<%=y%>"><% if x<>0 Then Response.write y2%></a> | 
        <a href="ord.asp?action=deliver&orderno=<%=orderno%>&x=<%=x%>"><% if y<>1  Then Response.write x2%><% if y=1 and x=0 Then Response.write x2%></a></font></td>             
 </tr>
      <%
rs.MoveNext
count1=count1+1
if trow="#FCFBFF" Then	   	   
trow="#FFFFFF"
Else
trow="#FCFBFF"
End if
loop 
End if %>
 <tr><td colspan="6">
<table border="0" width="100%" class="btntable">
        <tr>
          <td width="100%">
          &nbsp;
          <%if totalpages>1 Then %>
<a href="<%=url%>?pg=<%=Currentpage-1%>&amp;sid=<%=sid%>&amp;os=<%=os1%>"><%If currentpage>1 Then Response.write "<"%></a>
            (<%=currentpage&"-"&Totalpages%>) <a href="<%=url%>?pg=<%=Currentpage+1%>&amp;sid=<%=sid%>&amp;os=<%=os1%>"><%If currentpage<Totalpages Then Response.write ">"%></a></font>
            <%End if%>
            </td>
        </tr>
      </table>
</td></tr>
</table>
<%
End If 
%>
<%
 
'-----------------------------------------------------------------------------------
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="500">
  <tr>
    <td width="100%"><br>
        Selgitused :<br>
        <b>&nbsp;? </b>- kinnitamata tellimus;
&nbsp;<b>ok</b> - kinnitatud tellimus ;&nbsp;<b>*</b> - osaliselt tarnitud 
    tellimus<p>aktsepteeri - kinnita tellimus, kui raha on laekunud<br>
    tühista - tühista kinnitatud tellimus<br>
    tarni kaup - märgi tellimus tarnituks (suletud tellimused)<br>
    kaup
    tagasi - kui kinnitatud ja tarnitud tellimus tagastatakse </p>
    <p>&nbsp;</p>
    <p>&nbsp;</td>
  </tr>
</table>