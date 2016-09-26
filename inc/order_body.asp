<%

'----------------------------------------------------------------------------------------------------------
t=request.querystring("t")
pg=request.querystring("pg")
if pg="" Then pg=1

topage=Request.ServerVariables("HTTP_REFERER")
orn=request.querystring("orderno")
if orn<>"" Then 

sqlstmt = "SELECT * from orders  "
sqlstmt = sqlstmt & " WHERE orders.orderno = "& orn &" AND username='"&user&"' "
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
sqlstmt = sqlstmt & " WHERE orders.orderno = "& orn &" AND username='"&user&"' "
  
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
        <td valign="top"><br><%Response.write "<H1>Order : "&orn&"</H1>"%>
         <%="Date : "&adate%>
        <br><br>
        <%="Client :<b> " &customer &"</b>"%>
        <br><br>
        <%if len(descr)>1 then response.write "Notes : " & descr%><br>
        </td>
        <td>&nbsp;</td>
        <td valign="top">
        <p align="right"><img border="0" src="img/logo.gif"><br>
        <br>
        LN.COM<br>
        TEST<br>
        TEST<br>
        <a href="mailto:info@id31.net">info@LN.COM</a><br>
        <br>
&nbsp;</p>
        </td>
      </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td width="100%">
    <table border="0" class="maintable" cellpadding="0" cellspacing="0"  width="500">
      <tr>
        <td width="10" bgcolor="#F2F0FF" align="center"><b>&nbsp;?</b></td>
        <td width="300" bgcolor="#F2F0FF"><b>SKU</b>&nbsp;</td>
        <td width="20" bgcolor="#F2F0FF" align="center"><b>QTY</b></td>
        <td width="50" bgcolor="#F2F0FF" align="right"><b>Price</b></td>
        <td width="70" bgcolor="#F2F0FF" align="right"><b>Total</b></td>
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
        <td bgcolor="<%=trow%>"><a href="default.asp?sid=<%=sid%>&code=<% =code%>&"><%=code%></a><font color="#6699FF" title="<%=comment%>" size="1"><%=" " & left(comment,50)%></font></td>
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
        <td width="50%" align="right"><b><%=" Shipping : "%>
<%if sShipping >0 Then
    Response.write FormatNumber(sShipping,2) 
    Else
    Response.write "Total to pay"
    End if %></b></td>
      </tr>
      <tr>
        <td>
        <p align="right">&nbsp;</td>
        <td align="right"><b><%=" SubTotal : " &formatnumber(sTotal,2)%></b></td>
      </tr>
      <tr>
        <td>
        <p align="right">&nbsp;</td>
        <td align="right"><b><%="VAT 18% : " &FormatNumber(sVAT,2) %></b></td>
      </tr>
      <tr>
        <td>
        <p align="left">&nbsp;</td>
        <td align="right"><b><%= "Total : "& FormatNumber(stotalfin,2) %></b></td>
      </tr>
      <tr>
        <td>
        &nbsp;</td>
        <td align="right">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="2">
        Some information .....<p>&nbsp;&nbsp;<img src="img/arrow.gif" width="10" height="10">&nbsp;<a href="index.asp?sid=<%=sid%>&action=help&m=3#12">Warranty</a><br>
&nbsp;</td>
      </tr>
    </table>
    </td>
  </tr>
</table>
&nbsp;
<%
End If
End If
Else 
%>
  <% 
  

sqlstmt = "SELECT orders.orderno, Last(orders.customer) AS customer, Last(orders.adate) AS adate, Sum((orders.qty)*(orders.price)) AS sumoford, Last(orders.username) AS username , Last(orders.description) as Description, Last(orders.shipcode) AS shipcode"
sqlstmt = sqlstmt & " FROM orders "
sqlstmt = sqlstmt & " WHERE ((orders.username)='"&user&"') "
sqlstmt = sqlstmt & " GROUP BY orders.orderno "
if os=0 Then
os1=1
os2="Active orders list"
sqlstmt = sqlstmt & " HAVING (((Avg(orders.delivered))<1))" 
Else
os1=0
os2="Delivered orders list"
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
sum=(rs("sumoford"))
shipcode=rs("shipcode")
username=rs("username")
descr=rs("description")
if len(descr)>30 Then descr=Left(descr,30) & "..."
 %>
   <tr>
        <td bgcolor="<%=trow%>">&nbsp; 
    <img src="img/arrow.gif" width="10" height="10">&nbsp;<a href="order.asp?sid=<%=sid%>&os=<%=os%>&orderno=<%=orderno%>"><%=orderno%></a></td><td bgcolor="<%=trow%>"><%'=shipcode%></td>
<td bgcolor="<%=trow%>"><%=adate%></td><td bgcolor="<%=trow%>"><a href="mailto:<%=username%>"><%=customer%></a></td><td bgcolor="<%=trow%>"><%=Formatnumber(sum,2)%></td><td bgcolor="<%=trow%>"><font size="1"><%=descr%></font></td>             
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
%>
<% End if%>
<tr><td colspan="6">
<table border="0" width="100%" class="btntable">
        <tr><td width="100%">&nbsp;
          <%if totalpages>1 Then %>
          <a href="<%=url%>?pg=<%=Currentpage-1%>&amp;sid=<%=sid%>&amp;os=<%=os1%>"><%If currentpage>1 Then Response.write "<"%></a>
            (<%=currentpage&"-"&Totalpages%>) <a href="<%=url%>?pg=<%=Currentpage+1%>&amp;sid=<%=sid%>&amp;os=<%=os1%>"><%If currentpage<Totalpages Then Response.write ">"%></a></font>
            <%end if%>
            </td>
        </tr>
      </table>
</td></tr>

</table>

<%End If 
 
'-----------------------------------------------------------------------------------
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="500">
  <tr>
    <td width="100%"><br>
        Explanations :<br>
        <b>&nbsp;? </b>- open order;
&nbsp;<b>ok</b> - closed order ;&nbsp;<b>*</b> - partly delivered order<p>&nbsp;</td>
  </tr>
</table>