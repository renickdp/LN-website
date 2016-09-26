<%
if session("admin")<>1 Then response.redirect("default.asp")
'----------------------------------------------------------------------------------------------------------
sAction=request.querystring("action")
f=request.querystring("f")
if f="" Then f="l"

Select Case sAction

Case "products"%>
<h2>Tooteinfo haldus</h2>

<% 
topage=Request.ServerVariables("HTTP_REFERER")
action=request.querystring("action")
if request.form("pg") <>"" Then 
pg=request.form("pg")
Else
pg=request.querystring("pg")
End if
vat=request.cookies("vat")
code=request.querystring("code")
%>
<table><tr><td>
<form method="POST" name="groupform" action="adm_prod.asp?action=<%=action%>&amp;m=<%=m%>">
Kategooria&nbsp;:&nbsp;<select size="1" name="group" class="boxrequired" onchange="groupform.submit()">
     <option selected value="">---vali---</option>
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
 </form>
</td></tr></table>
<%if request.form("group")="" Then
group=request.querystring("group")
Else
group=request.form("group")
End if
yee=request.querystring("yee")
if yee="" Then yee=0

if yee<>1 Then
sqlstmt = "SELECT * from items  "
 if group<>"" Then sqlstmt = sqlstmt & " WHERE items.GROUP = '"&GROUP&"' "
Else
sqlstmt = "SELECT * from items WHERE yee=1 "
 if group<>"" Then sqlstmt = sqlstmt & " AND items.GROUP = '"&GROUP&"' "
End if

if  code<>"" Then
sqlstmt = "SELECT * from items  "
sqlstmt = sqlstmt & " WHERE items.code = '"&code&"' "

End if



rs.open sqlstmt, conn,1 ,1

Const NrPerPage = 20

If pg = "" then
    CurrentPage = 1 'We're on the first page
Else
    CurrentPage = CInt(pg)
End If
   
If rs.eof then
response.write "<center>Select a course from menu bar in the left"
response.write "<br></center>"
response.end
Else
  
If Not rs.EOF Then
 	rs.MoveFirst
 	rs.PageSize = NrPerPage
    TotalPages = rs.PageCount
    rs.AbsolutePage = CurrentPage

 %>
<table width="500" cellspacing="0" cellpadding="0" class="maintable" style="border-collapse: collapse" bordercolor="#111111">
<tr><th bgcolor="#ECECFF" colspan="5"><%=group%><font size="1">
  <a href="vat.asp"><% if vat=1 Then 
  Response.write "Hinnad käibemaksuta"
  Else
  Response.write "Hinnad käibemaksuga"
  End if%></a></font></th></tr>
<tr><td bgcolor="#ECECFF"><b>KOOD</b></td><td bgcolor="#ECECFF"><b>SELGITUS</b></td>
  <td bgcolor="#ECECFF" align="right"><b>Hinnakirjas </b></td>
  <td bgcolor="#ECECFF" align="right"><b>Omahind&nbsp;</b></td>
  <td bgcolor="#ECECFF" align="right"><b>Toimingud</b></td></tr>
<%
trow="#FCFBFF"
 Do while not rs.eof and Count1  < rs.PageSize
code = rs("code")
group = rs("group")
group1 = rs("group1")
group2 = rs("group2")
group3 = rs("group3")
product = rs("product")
prod=product
if len(product)>10 Then prod=left(product,10)&"..."
description=rs("description")
if request.querystring("code")="" Then description=left(description,20)&"<a href='default.asp?sid="&sid&"&code="&code&"'><img alt='Read more ...' border='0' src='img/s_arrow.gif'>" 
price_a=rs("price_a")
if price_a="" Then price_a=0
price_a=price_a*vat
price=rs("price")
if price="" Then price=0
'price=price-((price*discount)/100)
price=price*vat
image=rs("image")
if image="" Then image="no.gif"
notes=rs("notes")
if notes="" Then notes = Replace(notes, vbCrLf, "<br>")
doc=rs("doc")
yee=rs("yee")
if yee="" Then yee=0
if yee=0 Then 
y=1
y2="soodush."
Else
y=0
y2="tavah."
End If
if yee=1 then trow="#FFE7CE"%>
 <tr><td bgcolor="<%=trow%>" nowrap valign="top"><a href="<%=p_url%>?sid=<%=sid%>&code=<% =code%>&m=<%=m%>&amp;action=<%=sAction%>"><font size="1"><b><%=code%> &nbsp;</b></font></td>
   <td bgcolor="<%=trow%>" valign="top"><FONT title="<%=product%>" size="1"><%=description%></font></td>
   <td bgcolor="<%=trow%>" nowrap align="right" valign="top"><font size="1">
   <b><%=formatnumber(price,2)%></b></font></td>
   <td bgcolor="<%=trow%>" align="right" valign="top">&nbsp;<font size="1"><%=formatnumber(price_a,2)%>
   </font> </td><td bgcolor="#FFE7CE" align="right" valign="top">&nbsp;<font size="1"><a href="adm_prod.asp?action=specials&f=a&code=<%=code%>&description=<%=rs("description")%>">Avalehele</a>|<a href="prod.asp?action=yee&code=<%=code%>&y=<%=y%>"><%=y2%></a> </font> </td></tr>
<tr><td bgcolor="#ECECFF" colspan="5" height="2"></td></tr> 
<%if request.querystring("code")<>"" then %>
<tr><td nowrap valign="top">&nbsp;</td><td valign="top"><br>
  Tootja : <%=product%><br>
  Kategooria : <%=group&" - "&group1 %><br><br><font color="#c0C0C0" size="1"><%=notes%></font><br><br>
   <a title="Product description <%=code%>" target="_blank" href="doc/<%=doc%>"><% if len(doc)>4 then
  Response.write "<img border='0' alt='Datasheet' src='img/file_ico.gif'>&nbsp;"  &doc&"</a>"
  End If
  %></td>
   <td nowrap align="right" valign="top">
   <p align="center"><img alt="<%=description%>" border="0" src="img/items/<%=image%>"></td>
   <td align="right" valign="top"> 
   &nbsp;</td>
   <td align="right" valign="top"> 
   &nbsp;</td></tr>
<tr><th bgcolor="#ECECFF" colspan="5">&nbsp;</th></tr>
<%End if
rs.MoveNext
if trow="#FCFBFF" Then	   	   
trow="#FFFFFF"
Else
trow="#FCFBFF"
End if
count1=count1+1
loop 
if request.form("product")="" Then
product=request.querystring("product")
Else
product=request.form("product")
End if
if request.form("group")="" Then
group=request.querystring("group")
Else
group=request.form("group")
End if
group1=request.querystring("group1")

%>
<tr><td colspan="5">
          <form method="POST" name="pagenoform" action="<%=url%>?pg=<%=pg%>&amp;sid=<%=sid%>&amp;group=<%=group%>&amp;group1=<%=group1%>&amp;product=<%=product%>&amp;src=<%=src%>&amp;m=<%=m%>&amp;action=<%=sAction%>">
<table border="0" width="100%" class="btntable">
        <tr>
          <td width="100%">&nbsp;
          <%if totalpages>1 Then %>
                    <a href="<%=url%>?pg=<%=Currentpage-1%>&amp;sid=<%=sid%>&amp;group=<%=group%>&amp;group1=<%=group1%>&amp;product=<%=product%>&amp;src=<%=src%>&amp;m=<%=m%>&amp;action=<%=sAction%>"><%If currentpage>1 Then Response.write "<"%></a>
                <input type="hidden" name="page" value="1">
                <select class="boxrequired" size="1" onchange="pagenoform.submit()" name="pg">
                <%for pgs=1 to Totalpages%>
                <option <%if pgs=currentpage Then response.write " selected "%> value="<%=pgs%>"><%=pgs%></option>
                <%next%>
                </select> 
                <%'="  (" &currentpage&")" %>
                <%'=currentpage&"-"&Totalpages%> 
                 <a href="<%=url%>?pg=<%=Currentpage+1%>&amp;sid=<%=sid%>&amp;group=<%=group%>&amp;group1=<%=group1%>&amp;product=<%=product%>&amp;src=<%=src%>&amp;m=<%=m%>&amp;action=<%=sAction%>"><%If currentpage<Totalpages Then Response.write ">"%></a></font>
             <%End if%>
            </td>
        </tr>
      </table>
      </form>
</td></tr>
<%
End If
%>
</table>
<%
End if
rs.Close

Case "specials"
if f="l" Then
sqlstmt = "SELECT specials.code, specials.description, specials.startdate,specials.enddate, items.PRICE FROM items INNER JOIN specials ON items.CODE = specials.code "
sqlstmt = sqlstmt &  " WHERE (specials.startdate <= datevalue(date()) AND specials.enddate >= datevalue(date()) ) "
sqlstmt = sqlstmt & " ORDER BY specials.enddate desc" 
rs.open sqlstmt, conn,1 ,1
%>
<table border="0" cellpadding="0" cellspacing="0" class="maintable" width="100%">
<tr><th width="100%">Avalehe eripakkumised</th></tr>
  <tr><td width="100%">
<ul>
<%
If rs.eof then
response.write "<li><a href='adm_prod.asp?action=specials&f=a'> ... </a></li>"

else

Do while not rs.eof 
code = rs("code")
price = rs("price")
startdate=rs("startdate")
enddate=rs("enddate")
description=rs("description")
Response.write "<li><font size='1'><a href='" & mainpage &"?code="& server.urlencode(code) & "'>"&description & "</a><b> " & formatnumber(price,2) & "</b> "& startdate&"-"&enddate&"</font></li>"
Response.write "  <font color='red' size='1'><a href='prod.asp?action=delspecial&code="&code&"'>kustuta</a></font>"
rs.MoveNext
loop
rs.close
response.write "<li><a href='adm_prod.asp?action=specials&f=a'>Lisa ... </a></li>"
End If

%>
</ul>
</td>
    
  </tr>
</table>
<% 
End If

if f="a" Then
%>
<form method="POST" action="prod.asp?action=addspecial">
<table border="0" cellpadding="0" cellspacing="0" width="500">
  <tr>
    <td width="50%" align="right">Kirjeldus&nbsp;:&nbsp;</td>
    <td width="50%">
    <input type="text" name="description" size="30" class="boxrequired" value="<%=request.querystring("description")%>"></td>
  </tr>
  <tr>
    <td width="50%" align="right">Toode&nbsp;:&nbsp;
    </td>
    <td width="50%">
    <input type="text" name="code" size="20" class="boxrequired" value="<%=request.querystring("code")%>"></td>
  </tr>
  <tr>
    <td width="50%" align="right">Algus&nbsp;:&nbsp;</td>
    <td width="50%">
    <input type="text" name="startdate" size="12" class="boxrequired" value="<%=date()%>"></td>
  </tr>
  <tr>
    <td width="50%" align="right">Lõpp&nbsp;:&nbsp;</td>
    <td width="50%">
    <input type="text" name="enddate" size="12" class="boxrequired" value="<%=date()+15%>"></td>
  </tr>
  <tr>
    <td width="50%" align="right">&nbsp;</td>
    <td width="50%">&nbsp;</td>
  </tr>
  <tr>
    <td width="100%" align="center" colspan="2">
       <input type="image" src="img/btn_ok.gif" value="Submit" name="I1" border="0"></td>
  </tr>
</table>
 </form>
<%
End if
Case Else

response.write "Vali tegevus..."
End Select
%>