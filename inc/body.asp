<%

'----------------------------------------------------------------------------------------------------------
srt=request.querystring("srt")
if srt="" Then srt="code"
ord=request.querystring("ord")
if ord="" Then ord="asc"
if ord="asc" Then
ord1="desc"
Else
ord1="asc"
End if

src=request.form("src")
if src="" Then src=request.querystring("src")
if src<>"" Then 
src= Replace(src, chr(39),"''")
if request.form("src")<>"" Then response.redirect(mainpage&"?sid="&sid&"&src="&src&"&m="&m)
End If
t=request.querystring("t")
qty=Session("qty")
if t="" Then t=0
topage=Request.ServerVariables("HTTP_REFERER")
action=request.querystring("action")
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
group2=request.querystring("group2")
group3=request.querystring("group3")
group= Replace(group, chr(39),"''")
group1= Replace(group1, chr(39),"''")
group2= Replace(group2, chr(39),"''")
group3= Replace(group3, chr(39),"''")
code=request.querystring("code")
product= Replace(product, chr(39),"''")
if request.form("pg") <>"" Then 
pg=request.form("pg")
Else
pg=request.querystring("pg")
End if
if pg="" Then pg=1


' --------------------------------------------product list -----------------------------------------------
if group<>"" or product<>"" or code<>"" or src<>"" Then 
sqlstmt = "SELECT * from items  "

if group<>"" or code<>"" and src="" Then 
if  code="" Then
if group<>"" Then  sqlstmt = sqlstmt & " WHERE items.GROUP = '"&GROUP&"' "
if group1<>"" Then sqlstmt = sqlstmt & " AND items.GROUP1 = '"&GROUP1&"' "
if group2<>"" Then sqlstmt = sqlstmt & " AND items.GROUP2 = '"&GROUP2&"' "
if group3<>"" Then sqlstmt = sqlstmt & " AND items.GROUP3 = '"&GROUP3&"' "
Else
sqlstmt = sqlstmt & " WHERE items.code = '"&code&"' "
End if

Else
sqlstmt = sqlstmt & " WHERE items.code Like '%"&src&"%' or items.GROUP Like '%"&src&"%' or items.GROUP1 Like '%"&src&"%' or items.GROUP2 Like '%"&src&"%' or items.GROUP3 Like '%"&src&"%' or items.description Like '%"&src&"%'  "
End if

sqlstmt = sqlstmt & " ORDER by items."&srt & " " & ord

rs.open sqlstmt, conn,1 ,1

Const NrPerPage = 20

If pg = "" then
    CurrentPage = 1 'We're on the first page
Else
    CurrentPage = CInt(pg)
End If
   
If rs.eof then
If src="" Then
response.write "<center>Select a course from menu bar in the left"
response.write "<br></center>"
Else
%>
<table class="maintable" width="500" border="0" style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0"><tr>
   <th colspan="2">Results :</th></tr><tr><td colspan="2">&nbsp;</td></tr>
  <tr>
    <td width="100%" colspan="2">
    No matches. Search again?</td>
  </tr>
  <tr><td align="right" width="50%">
    &nbsp;</td>
    <td align="left" width="50%">
    &nbsp;</td></tr>
  <tr><td align="right" width="50%">
    <a href="index.asp?action=search">Yes</a>&nbsp;&nbsp;</td>
    <td align="left" width="50%">
    &nbsp;&nbsp; <a href="default.asp">No</a></td></tr>
  <tr><td align="center" colspan="2"></td>
   </tr></table>
<%
End If
response.end
Else
  
If Not rs.EOF Then
 	rs.MoveFirst
 	rs.PageSize = NrPerPage
    TotalPages = rs.PageCount
    rs.AbsolutePage = CurrentPage

 %>
<table width="500" cellspacing="0" cellpadding="0" class="maintable" border="0">
<tr><th bgcolor="#FFE7CE" colspan="4">
<%
if group<>"" Then response.write group& " "
if group1<>"" And group1<>Cstr("---") Then response.write " : " & group1
if group2<>"" And group2<>Cstr("---") Then response.write " : " & group2
if group3<>"" And group3<>Cstr("---") Then response.write " : " & group3
if product<>"" Then response.write product& " "
if code<>"" Then response.write code 
if src<>"" Then response.write "Otsitud : " &src 
 %></th></tr>
<tr><td bgcolor="#FFE7CE">
  &nbsp;<img border="0" src="img/<%=ord%>.gif">
  <a title="Sorteeri" href="<%=mainpage%>?sid=<%=sid%>&pg=<%=pg%>&group=<%=group%>&group1=<%=group1%>&group2=<%=group2%>&group3=<%=group3%>&product=<%=product%>&m=<%=m%>&srt=code&ord=<%=ord1%>"><%="<b>SKU</b>"%></a></td><td bgcolor="#FFE7CE">
  <b>Description</b></td>
  <td bgcolor="#FFE7CE" align="right"><b>Price <font size="1">
  <a href="vat.asp"><% if vat=1 Then 
  Response.write "(+ VAT)"
  Else
  Response.write "(VAT inc.)"
  End if%></a></font></b></td>
  <td bgcolor="#FFE7CE" align="right"> 
   &nbsp;</td></tr>
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
if request.querystring("code")="" Then description=left(description,30)&"<a href='index.asp?sid="&sid&"&code="&code&"'><img alt='Read more ...' border='0' src='img/s_arrow.gif'>" 
price=rs("price")
price=price-((price*discount)/100)
price=price*vat
image=rs("image")
if image="" Then image="no.gif"
notes=rs("notes")
if notes<>"" Then 
notes=LinkURLS(notes)
notes = Replace(notes, vbCrLf, "<br>")
End If
doc=rs("doc")
yee=rs("yee")
if yee="" Then yee=0
if yee=1 then trow="#FFE7CE"%>
 <tr><td bgcolor="<%=trow%>" nowrap valign="top"><a href="<%=p_url%>?sid=<%=sid%>&code=<% =server.Urlencode(code)%>&m=<%=m%>"><font size="1"><b><%=code%> &nbsp;</b></font></td>
   <td bgcolor="<%=trow%>" valign="top"><FONT title="<%=product%>" size="1"><%="<b>"&prod&"</b> : "& description%></font></td>
   <td bgcolor="<%=trow%>" nowrap align="right" valign="top"><font size="1">
   <b><%=formatnumber(price,2)%></b></font></td>
   <td bgcolor="<%=trow%>" align="right" valign="top"> 
   &nbsp;<a href="?sid=<%=sid%>&action=add&item=<% =rs("code")%>&count=1&t=1"><img border="0" src="img/add.gif" alt="Lisa ostukorvi"></a><font size="1">
   </font> </td></tr>
<tr><td bgcolor="#FFE7CE" colspan="4" height="2"></td></tr> 
<%if request.querystring("code")<>"" then %>
<tr><td nowrap valign="top">&nbsp;</td><td valign="top"><br>
  Brand : <%=product%><br>
  Category : <%=group&" - "&group1 %><br><br><font color="#c0C0C0" size="1"><%=notes%></font><br><br>
   <a title="Product description <%=code%>" target="_blank" href="doc/<%=doc%>"><% if len(doc)>4 then
  Response.write "<img border='0' alt='Datasheet' src='img/file_ico.gif'>&nbsp;"  &doc&"</a>"
  End If
  %></td>
   <td nowrap align="right" valign="top">
   <p align="center"><img alt="<%=description%>" border="0" src="img/items/<%=image%>"></td>
   <td align="right" valign="top"> 
   &nbsp;</td></tr>
<tr><th bgcolor="#FFE7CE" colspan="4">&nbsp;</th></tr>
<%End if
rs.MoveNext
if trow="#FCFBFF" Then	   	   
trow="#FFFFFF"
Else
trow="#FCFBFF"
End if
count1=count1+1
loop 
group1=request.querystring("group1")
group2=request.querystring("group2")
group3=request.querystring("group3")
group= Replace(group, chr(39),"''")
group1= Replace(group1, chr(39),"''")
group2= Replace(group2, chr(39),"''")
group3= Replace(group3, chr(39),"''")
code=request.querystring("code")
product= Replace(product, chr(39),"''")
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
<tr><td colspan="4">
          <form method="POST" name="pagenoform" action="<%=url%>?pg=<%=pg%>&amp;sid=<%=sid%>&amp;group=<%=group%>&amp;group1=<%=group1%>&amp;group2=<%=group2%>&amp;group3=<%=group3%>&amp;product=<%=product%>&amp;src=<%=src%>&amp;m=<%=m%>">
<table border="0" width="100%" class="btntable">
        <tr>
          <td width="100%">&nbsp;
          <%if totalpages>1 Then %>
                    <a href="<%=url%>?pg=<%=Currentpage-1%>&amp;sid=<%=sid%>&amp;group=<%=group%>&amp;group1=<%=group1%>&amp;group2=<%=group2%>&amp;group3=<%=group3%>&amp;product=<%=product%>&amp;src=<%=src%>&amp;m=<%=m%>"><%If currentpage>1 Then Response.write "<"%></a>
                <input type="hidden" name="page" value="1">
                <select class="boxrequired" size="1" onchange="pagenoform.submit()" name="pg">
                <%for pgs=1 to Totalpages%>
                <option <%if pgs=currentpage Then response.write " selected "%> value="<%=pgs%>"><%=pgs%></option>
                <%next%>
                </select> 
                <%'="  (" &currentpage&")" %>
                <%'=currentpage&"-"&Totalpages%> 
                 <a href="<%=url%>?pg=<%=Currentpage+1%>&amp;sid=<%=sid%>&amp;group=<%=group%>&amp;group1=<%=group1%>&amp;group2=<%=group2%>&amp;group3=<%=group3%>&amp;product=<%=product%>&amp;src=<%=src%>&amp;m=<%=m%>"><%If currentpage<Totalpages Then Response.write ">"%></a>
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
Else 
' --------------------------------------------specials -----------------------------------------------
if action="" Then 
sqlstmt = "SELECT * from items  WHERE items.yee = 1 "
 sqlstmt = sqlstmt & " ORDER by items.code"
 rs.open sqlstmt, conn,1 ,1
   
If rs.eof then
response.write "<center>Select a course from menu bar in the left"
response.write "<br></center>"
response.end
Else
  
If Not rs.EOF Then
 	rs.MoveFirst
 %>

<table width="500" cellspacing="0" cellpadding="0" class="maintable" border="0" style="border-collapse: collapse" bordercolor="#111111">
<tr><th bgcolor="#FFE7CE" colspan="4">We Recommend</th></tr>
<tr><td align="right" colspan="4" height="2"><a href="vat.asp">
<% if vat=1 Then 
  Response.write "<font size='1'>(* Hinnad käibemaksuta)</font>"
  Else
  Response.write "<font size='1'>(* Hinnad käibemaksuga)</font>"
  End if%></a>
</td></tr>

<tr><td bgcolor="#FFE7CE">&nbsp;</td><td bgcolor="#FFE7CE">&nbsp;</td>
  <td bgcolor="#FFE7CE" align="right">&nbsp;</td>
  <td bgcolor="#FFE7CE" align="right"></td></tr>
<%
trow="#FCFBFF"
 Do while not rs.eof 
code = rs("code")
grp=rs("group")
grp1=rs("group1")
product = rs("product")
description=rs("description")
price=rs("price")
price=price-((price*discount)/100)
price=price*vat
image=rs("image")
if image="" Then image="no.gif"
notes=rs("notes")
if notes<>"" Then 
notes=LinkURLS(notes)
notes = Replace(notes, vbCrLf, "<br>")
End If
 %>
 <tr><td bgcolor="<%=trow%>" nowrap valign="top"></td>
   <td bgcolor="<%=trow%>" valign="top"><FONT title="<%=product%>" size="1"><b><%=product &" - "& grp&" - "& grp1%></b><br><a href="?sid=<%=sid%>&code=<% =code%>&"><%=description%></a><br><font color="#6699FF" size="1"><%=notes%></font></font></td>
   <td bgcolor="<%=trow%>" nowrap align="right" valign="top"><font size="1"><br><a href="?sid=<%=sid%>&code=<% =code%>&"><img alt="<%=description%>" border="0" src="img/items/<%=image%>"></a><br><b><font size="3" color="CC3300"><%=formatnumber(price,2)%></font></b></font>
   <a href="?sid=<%=sid%>&action=add&item=<% =rs("code")%>&count=1&t=1"><img border="0" src="img/add.gif" alt="Buy this right now"></a>
   </td>
   <td bgcolor="<%=trow%>" align="right" valign="top"> &nbsp; </td></tr>
<tr><td bgcolor="#FFE7CE" colspan="4" height="2"></td></tr> 
<%
rs.MoveNext
if trow="#FCFBFF" Then	   	   
trow="#FFFFFF"
Else
trow="#FCFBFF"
End if

loop 
rs.close
End If
%>
</table>
<%End If
End if
End If
'-----------------------------------------------------------------------------------
%>
<% 
Sub ChangeItemInCart(iItemID, iItemCount)
 	
 	If dictCart.Exists(iItemID) Then
		if iItemCount<=0 Then 
			dictCart.Remove iItemID
		Else
		dictCart(iItemID) = iItemCount
		End If
	Else
		dictCart.Add iItemID, iItemCount
	End If

call CountCart()
End Sub


Sub AddItemToCart(iItemID, iItemCount)
 	If dictCart.Exists(iItemID) Then
		dictCart(iItemID) = dictCart(iItemID) + iItemCount
	Else
		dictCart.Add iItemID, iItemCount
	End If
	call countcart()
End Sub

Sub RemoveItemFromCart(iItemID, iItemCount)
 	If dictCart.Exists(iItemID) Then
		If dictCart(iItemID) <= iItemCount Then
			dictCart.Remove iItemID
		Else
			dictCart(iItemID) = dictCart(iItemID) - iItemCount
		End If
	Else
		Response.Write "Couldn't find any of that item your cart.<BR><BR>" & vbCrLf
	End If
	call countcart()
End Sub

Sub ShowItemsInCart()
Dim Key
Dim aParameters ' as Variant (Array)
Dim sTotal, sShipping	
if qty>0 Then	%>
<TABLE Border="0"  class="maintable" CellPadding="2" CellSpacing="0" width="500">
  <TR align="center"> 
    <th bgcolor="#FFE7CE" colspan="6">
      Your cart</th>
  </TR>
  <TR align="center"> 
    <TD bgcolor="#E6E6E6"><b>&nbsp;</b></TD>
    <TD bgcolor="#E6E6E6"><b>&nbsp;SKU</b></TD>
    <TD bgcolor="#E6E6E6"><b>Descr</b></TD>
    <TD bgcolor="#E6E6E6"><b>&nbsp;QTY</b></TD>
    <TD bgcolor="#E6E6E6"><b>&nbsp;Price</b></TD>
    <TD align="justify" bgcolor="#E6E6E6">
    <p align="right"><b>Total</b></TD>
  </TR>
  <%
	sTotal = 0
		For Each Key in dictCart
	aParameters = GetItemParameters(Key)
		%> 
  <TR> <form method="POST" name="<%=a%>" action="<%=mainpage%>?action=change&m=<%=m%>" >
<input type="hidden" name="code" value="<%=key%>">
    <TD align="center" bgcolor="#FFE7CE" nowrap>
    &nbsp;<a href="<%=mainpage%>?sid=<%=sid%>&action=del&item=<%= Key %>&count=<%= dictCart(Key) %>">x</a>&nbsp;<a href="<%=mainpage%>?sid=<%=sid%>&action=add&item=<%= Key %>&count=1"> </a></TD>
    <TD nowrap><b><a href="?sid=<%=sid%>&code=<% =key%>&"><%= Key %></a></b></TD>
    <TD><font size="1"><%= left(aParameters(1),20)%></font> </TD> 
    <TD align="center" nowrap>     <a href="<%=mainpage%>?sid=<%=sid%>&action=del&item=<%= Key %>&count=1">-</a> <b>
    &nbsp;<input type="text" name="qty" value="<%= dictCart(Key) %>" size="3" class="boxrequired">&nbsp;</b> <a href="<%=mainpage%>?sid=<%=sid%>&action=add&item=<%= Key %>&count=1">+</a></TD>
    <TD align="right" nowrap><%= Formatnumber(aParameters(2),2) %></TD>
    <TD ALIGN="right" nowrap><b>&nbsp;<%= FormatNumber(dictCart(Key) * CSng(aParameters(2)),2) %></b>&nbsp;<input type="image" src="img/btn-ok.gif" alt="Muuda" value="Submit" name="OK" border="0"></TD>
</form>
  </TR>
  <% 
		sTotal = sTotal + (dictCart(Key) * CSng(aParameters(2)))
		Session("Total")=sTotal
	Next
		'
	If sTotal <> 0 and sTotal < 2000 Then
		sShipping = 0
		sShipping = 65
	Else
		sShipping = 0
	End If
	If request.form("shippment")=2 then sShipping=0
	Session("Shipping")=sShipping
	sTotal = sTotal + sShipping
	sVAT = (sTotal) * 0.18
	stotalfin=stotal+svat
	%> 
  <TR> 
    <TD COLSPAN=5 ALIGN="right" bgcolor="#FFE7CE" height="2"></TD>
    <TD ALIGN="right" nowrap bgcolor="#FFE7CE" height="2"></TD>
  </TR>
  <TR> 
    <TD COLSPAN=5 ALIGN="right" bgcolor="#E6E6E6"><b>Shipping 
    <a title="Tellimused, mille väärtus on üle 2000 krooni, pakume tasuta kohaletoimetamist Eesti piires " href="javascript:void(0)">
    *</a> :</b></TD>
    <TD ALIGN="right" nowrap bgcolor="#E6E6E6"><b>&nbsp;&nbsp;
    <%if sShipping >0 Then
    Response.write FormatNumber(sShipping,2) 
    Else
    Response.write "Tasuta"
    End if %></b></TD>
  </TR>
  <tr>
    <TD COLSPAN=5 ALIGN="right" bgcolor="#E6E6E6"><b>Subtotal :</b></TD>
    <TD ALIGN="right" nowrap bgcolor="#E6E6E6"><b>&nbsp;&nbsp;<%= FormatNumber(sTotal,2) %></b></TD>
    </tr>
  <TR> 
    <TD COLSPAN=5 ALIGN="right" bgcolor="#E6E6E6"><b>VAT 18 % :</b></TD>
    <TD ALIGN="right" nowrap bgcolor="#E6E6E6"><b>&nbsp;&nbsp;<%= FormatNumber(sVAT,2) %></b></TD>
    </tr>
  <TR> 
    <TD COLSPAN=5 ALIGN="right" bgcolor="#FFE7CE"><b>Total :</b></TD>
    <TD ALIGN="right" nowrap bgcolor="#FFE7CE"><b>&nbsp;&nbsp;<%= FormatNumber(stotalfin,2) %></b></TD>
    </tr>
</TABLE>

     	<%Else
		Response.write "<table class='maintable'><tr><th width='500'> Kott on tühi ...</th></tr></table>"
		End If
		call countcart()
	End Sub
%>
 
<%

Function GetItemParameters(iItemID)
SQL = "SELECT * FROM Items WHERE code like '%" & iItemID & "%'"
rs.Open SQL, Conn
rs.MoveFirst

Dim aParameters
price=rs("Price")
descr=rs("description")
prod=rs("Product")
price_a=rs("price_a")
if price_a="" Then price_a=0
price_a=price_a*vat
price=price-((price*discount)/100)
aParameters = Array(prod, descr, price, price_a)
GetItemParameters = aParameters
rs.close
End Function

%>

<% 
Dim dictCart ' as dictionary
Dim sAction ' as string
Dim iItemID ' 
Dim iItemCount ' as integer

If IsObject(Session("cart")) Then
	Set dictCart = Session("cart")
Else

	Set dictCart = Server.CreateObject("Scripting.Dictionary")
End If

sAction = CStr(Request.QueryString("action"))
sStep = CStr(Request.QueryString("step"))
iItemID = CStr(Request.QueryString("item"))
iItemCount = CInt(Request.QueryString("count"))

%>
       
 <%
Select Case sAction
	Case "add"
		
	AddItemToCart iItemID, iItemCount
		ShowItemsInCart
		
		Response.redirect(topage)
	if qty>0 Then %> 

   <% End If
   
   Case "change"	
iItemID=request.form("code")
iItemCount=request.form("qty")
if iItemCount="" Then iItemCount=0
	ChangeItemInCart iItemID, iItemCount
		ShowItemsInCart
		
		Response.redirect(topage)


	Case "del"
	
	RemoveItemFromCart iItemID, iItemCount
	ShowItemsInCart
	Response.redirect(topage)
	if qty>0 Then %> 
   <% End if
	Case "viewcart"
				ShowItemsInCart
		if qty>0 Then%>
<TABLE class="maintable" BORDER=0 CELLSPACING=0 CELLPADDING=0 width="500" >       

 <TR> <TD align="center">
     <br>
   <%if Session("Total") >= 1000 Then%>
   &nbsp;<img src="img/arrow.gif" width="10" height="10">&nbsp;
   <a href="<%=mainpage%>?sid=<%=sid%>&action=login">Submit order </a> 
	<%Else%> Minimum order valuemust be&nbsp; 1000 .
   <%End if%>
         </td></tr></table>                
 <% End If
case "checkout"
ShowItemsInCart
if Session("logonid") = "" Then Response.redirect(mainpage&"?sid="& sid & "&action=login")
%>       

<% if qty>0 Then%>

<table border="0" cellpadding="0" cellspacing="0" width="500" class="maintable">
  <tr>
    <th width="100%" colspan="2" height="19">Checkout</th>
  </tr>
  <tr>
    <td width="100%" align="center" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td width="100%" align="center" colspan="2"><font size="1">This is an 
    agreement between&nbsp; <%=fullname%> and ... blabla</font></td>
  </tr>
  <tr>
    <td width="100%" align="center" colspan="2">
    <p align="center">&nbsp;</td>
  </tr>
  <form method="POST" name="shippmentform" action="<%=mainpage&"?sid="&sid & "&action=checkout"%>">  
  <%smt=request.form("shippment")%>
  <td width="50%" align="right" >Shipping&nbsp;:&nbsp; </td>
    <td width="50%"> 
    <select size="1" name="shippment" class="boxrequired" onchange="shippmentform.submit()">
<% if smt="" Then %>   
 <option selected>--pick---</option>
    <option value="2">Postal service</option>
    <option value="1">UPS</option>
<%Else %>
    <option <%if smt=2 Then Response.write "selected" %> value="2">Postal service</option>
    <option <%if smt=1 Then Response.write "selected" %> value="1">UPS</option>
<%End if %> 
   </select></td>
  </tr>
  </form>
  <tr>
    <td width="50%" align="right">&nbsp;</td>
    <td width="50%">&nbsp;</td>
  </tr>
  <% if smt<>"" Then %>
  <tr>
    <td width="50%" align="right" >Payment&nbsp;:&nbsp;</td>
    <td width="50%">&nbsp;Credit Card</td>
  </tr>
  
 <%'rs.close
SQL = "SELECT* FROM users WHERE  login='" & user & "' "
rs.Open SQL, Conn

Do while not rs.eof 
coaddress = rs("coaddress")
coname=rs("coname")
phone=rs("phone")
description=rs("fullname")
%>
<form method="POST" name="cart" action="co.asp?sid=<%=sid%>&action=finishcart">
  <tr>
    <td width="50%" align="right" height="19">&nbsp;</td>
    <td width="50%"> &nbsp;</td>
  </tr>
  <tr>
    <th width="100%" align="right" colspan="2">
    <p align="center">Delivery address</th>
  </tr>
  <tr>
    <td width="50%" align="right">Name :&nbsp; </td>
    <td width="50%"> <input type="text" name="fullname" class="boxrequired" size="25" value="<%=coname%>">&nbsp;</td>
  </tr>
  <tr>
    <td width="50%" align="right">Adress :&nbsp; </td>
    <td width="50%"><input type="text" name="coaddress" size="30" class="boxrequired" value="<%=coaddress%>">&nbsp;</td>
  </tr>
  <tr>
    <td width="50%" align="right" >Phone :&nbsp; </td>
    <td width="50%"><input type="text" name="phone" size="10" class="boxrequired" value="<%=phone%>">&nbsp;</td>
  </tr>
  <tr>
    <td width="50%" align="right">Notes :&nbsp; </td>
    <td width="50%"><input type="text" name="notes" size="30" class="boxrequired" value="<%=description%>">&nbsp;</td>
  </tr>
  <tr>
     <td colspan="2">
&nbsp;</td>
  </tr>
   <%
 rs.MoveNext
 loop
 rs.close
%> 
  <tr> <td colspan="2">
       <p align="center">
       <input type="image" src="img/btn_ok.gif" value="Submit" name="B1" border="0"> </p>
          <input type="hidden" name="user" value="<%=user%>">
         <input type="hidden" name="smt" value="<%=smt%>">
                  <input type="hidden" name="shipping" value="<%=session("Shipping")%>">
  
     </td>
  </tr>
  </form>
  <%End if%>
</table>


<%End if%> 
   
 <% case "search"

%>
<table class="maintable" width="500" border="0" cellpadding="0" cellspacing="0"><tr>
  <form method="POST" name="searchform" action="<%="index.asp?sid="&sid  %>">
  <th colspan="2">Search</th></tr><tr><td colspan="2">&nbsp;</td></tr>
  <tr>
    <td width="50%">
    <p align="right">SKU/category&nbsp; :&nbsp; </td>
    <td width="50%">
         <p>
         <input type="text" name="src" size="15" class="boxrequired" value="<%=request.form("src")%>"></p>
     </td>
  </tr>
  <tr><td>&nbsp;</td><td>&nbsp;</td></tr>
  <tr><td colspan="2">
    <p align="center">
       <input type="image" src="img/btn_ok.gif" value="Submit" name="I1" border="0"></p></td></form>
    </tr></table>
<% case "help"%>
<!--#INCLUDE FILE="help.asp"-->        
<%
Case "login" 
' -----------------------------------------------------------------------------------------
%>
<%if Session("logonid") = ""  and l< 4 Then %>
<table border="0" cellpadding="0" cellspacing="0"  width="500" class="maintable">
    <form method="POST" name="loginform" action="login.asp">
  <INPUT type="hidden" name="dologin" value="yes">
<%    	sURL = Request.ServerVariables("SCRIPT_NAME")
    	if Request.ServerVariables("QUERY_STRING") <> "" Then
    	sURL = sURL & "?" & Request.ServerVariables("QUERY_STRING")
   	End if%>
  <INPUT type="hidden" name="siteid" value="<%=surl%>">
<tr>
    <th colspan="2" width="50%">Log in </th></tr>
  <tr>
    <td width="50%">
    &nbsp;</td>
    <td width="50%">
         &nbsp;</td>
  </tr>
  <tr>
    <td align="right" width="50%">
    User id :&nbsp; </td>
    <td width="50%">
         <p>
         <input type="text" name="id" size="10" class="boxrequired" value="<%=request.querystring("login")%>"></p>
     </td>
  </tr>
  <tr>
    <td width="50%">
    <p align="right">password :&nbsp; </td>
    <td width="50%">
      <input type="password" name="pwd" size="10" class="boxrequired"></td>
  </tr>
  <tr>
       <td align="center" colspan="2">&nbsp;</td>
  </tr>
  <tr>
       <td align="center" colspan="2"> <input type="image" src="img/btn_ok.gif" value="Submit" name="I1" border="0"></td>
  </tr>
  <tr>
    <td colspan="2" width="50%">
    <p align="center"><br>
    <font face="Verdana, Arial, Helvetica, sans-serif" size="2">
   &nbsp;<img src="img/arrow.gif" width="10" height="10">&nbsp; </font>
         <a href="<%=mainpage%>?sid=<%=sid%>&action=register&m=3">If you do not 
    have registered yet click here.</a> </td></tr>
  </form>
</table>
<% else
If l>3 Then %>

    <table class="maintable" width="500" border="0" style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0"><tr>
      <th colspan="2">Cannot remember password ? ...</th></tr>
      <tr>
      <form method="POST" name="reminder" action="remind.asp">
    <td width="50%" align="right">&nbsp;</td>
    <td width="50%">
    &nbsp;</td>
      </tr>
      <tr>
    <td width="100%" align="center" colspan="2">We will remind it to you </td>
      </tr>
      <tr>
    <td width="50%" align="right">Your email address :&nbsp; </td>
    <td width="50%">
    <input type="text" name="remind" size="15" class="boxrequired" value="<%=request.form("login")%>"> </td>
      </tr>
      <tr><td>&nbsp;</td><td>&nbsp;</td></tr><tr><td align="center" colspan="2"><input type="image" src="img/btn_ok.gif" value="Submit" border="0" name="I1"></td></tr></form></table>

    <%
    Else
    if qty>0 Then Response.Redirect(mainpage&"?sid="& sid & "&action=checkout")
	Response.Redirect("default.asp")
	End if

End If%>
<%
Case "register" ' 
'---------------------------------------------------------------------------------------------
%>
<table border="0" cellpadding="0" cellspacing="0" width="500" class="maintable">
  <form method="POST" name="reg" action="reg.asp?action=register">
  <tr>
    <th width="100%" colspan="2">Register</th>
  </tr>
  <tr>
    <td width="50%" align="right">&nbsp;</td>
    <td width="50%">
    &nbsp;</td>
  </tr>
  <tr>
    <td width="50%" align="right">email adress :&nbsp; </td>
    <td width="50%">
    <input type="text" name="login" size="15" class="boxrequired" value="<%=request.form("login")%>"> </td>
  </tr>
  <tr>
    <td width="50%" align="right">Full name :&nbsp; </td>
    <td width="50%">
    <input type="text" name="fullname" size="25" class="boxrequired" value="<%=request.form("fullname")%>"></td>
  </tr>
  <tr>
    <td width="50%" align="right">Company name :&nbsp; </td>
    <td width="50%">
    <input type="text" name="coname" size="25" class="boxrequired" value="<%=request.form("coname")%>"></td>
  </tr>
  <tr>
    <td width="50%" align="right">Address :&nbsp; </td>
    <td width="50%">
    <input type="text" name="coaddress" size="30" class="boxrequired" value="<%=request.form("coaddress")%>"></td>
  </tr>
  <tr>
    <td width="50%" align="right">Phone :&nbsp; </td>
    <td width="50%">
    <input type="text" name="phone" size="10" class="boxrequired" value="<%=request.form("phone")%>"></td>
  </tr>
  <tr>
    <td width="50%" align="right">Password :&nbsp; </td>
    <td width="50%">
    <input type="password" name="pass1" size="10" class="boxrequired"></td>
  </tr>
  <tr>
    <td width="50%" align="right">Retype password :&nbsp; </td>
    <td width="50%">
    <input type="password" name="pass2" size="10" class="boxrequired"></td>
  </tr>
  <tr>
    <td width="100%" align="right" colspan="2">
    <p align="center"><br>
    PLEASE FILL ALL FIELDS !</td>
  </tr>
  <tr>
    <td width="100%" align="right" colspan="2">
    &nbsp;</td>
  </tr>
  <tr>
    <td width="100%" align="center" colspan="2">
    <input type="image" src="img/btn_ok.gif" value="Submit" border="0" name="I1"></td>
  </tr>
</table>



<% Case "account" %>

<table border="0" cellpadding="0" cellspacing="0" width="500" class="maintable">
  <tr>
    <th width="100%" colspan="2">Your Data</th>
  </tr>
  <tr>
    <td width="100%" align="center" colspan="2">&nbsp;</td>
  </tr>
    
 <%'rs.close
SQL = "SELECT* FROM users WHERE  login='" & user & "' "
rs.Open SQL, Conn

Do while not rs.eof 
id=rs("ID")
pass=rs("pass")
login=rs("login")
fullname=rs("fullname")
coaddress = rs("coaddress")
coname=rs("coname")
phone=rs("phone")
description=rs("fullname")

if request.querystring("step")<>"pwd" Then%>
<form method="POST" name="accountform" action="reg.asp?sid=<%=sid%>&action=change">
  <tr>
    <td width="50%" align="right">email/ username :&nbsp; </td>
    <td width="50%"><%=login%></td>
  </tr>
  <tr>
    <td width="50%" align="right">Full name :&nbsp; </td>
    <td width="50%">
    <input type="text" name="fullname" size="25" class="boxrequired" value="<%=fullname%>"></td>
  </tr>
  <tr>
    <td width="50%" align="right">Company name :&nbsp; </td>
    <td width="50%">
    <input type="text" name="coname" size="25" class="boxrequired" value="<%=coname%>"></td>
  </tr>
  <tr>
    <td width="50%" align="right">Address :&nbsp; </td>
    <td width="50%">
    <input type="text" name="coaddress" size="30" class="boxrequired" value="<%=coaddress%>"></td>
  </tr>
  <tr>
    <td width="50%" align="right">Phone :&nbsp; </td>
    <td width="50%">
    <input type="text" name="phone" size="10" class="boxrequired" value="<%=phone%>"></td>
  </tr>
    <tr>
    <td width="100%" align="right" colspan="2">    &nbsp;</td>
  </tr>
    <tr>
    <td width="100%" align="center" colspan="2">    
    <input type="image" src="img/btn_ok.gif" value="Submit" name="I1" border="0"></td>
  </tr>
    <tr>
    <td width="100%" align="right" colspan="2">    &nbsp;</td>
    <input type="hidden" name="id" value="<%=id%>">
  </tr>
  <tr> 
  <td colspan="2">  <p align="center">
  <a href="<%=mainpage%>?sid=<%=sid%>&action=account&step=pwd&m=3">Change your 
  password</a></p>
             </td>
  </tr>
   <% End if 
 rs.MoveNext
 loop
 rs.close
%> </form>
  
  <tr> <td colspan="2">
       &nbsp;</td>
  </tr>
<%if request.querystring("step")="pwd" Then%> 
 <form method="POST" name="cangepwdform" action="reg.asp?sid=<%=sid%>&action=pwd">
    <tr>
    <td width="50%" align="right">password :&nbsp; </td>
    <td width="50%">
    <input type="password" name="pass0" size="10" class="boxrequired"></td>
  </tr>

  <tr>
    <td width="50%" align="right">New password :&nbsp; </td>
    <td width="50%">
    <input type="password" name="pass1" size="10" class="boxrequired"></td>
  </tr>
  <tr>
    <td width="50%" align="right">Retype new password :&nbsp; </td>
    <td width="50%">
    <input type="password" name="pass2" size="10" class="boxrequired"></td>
  </tr>
    <tr>
  <td width="100%" align="right" colspan="2">    
    &nbsp;</td>
  </tr>
    <td width="100%" align="center" colspan="2">    
    <input type="image" src="img/btn_ok.gif" value="Submit" name="I1" border="0"></td>
  </tr>
    <tr> 
  <td colspan="2">  &nbsp;</td>
  </tr>
    <tr> 
  <td colspan="2">  <p align="center">
  <a href="<%=mainpage%>?sid=<%=sid%>&action=account&m=3&step=">Change my data</a></p>
             </td>
  </tr>
    <input type="hidden" name="pass" value="<%=pass%>">
    <input type="hidden" name="id" value="<%=id%>">
    <input type="hidden" name="login" value="<%=login%>">
  </form>
<%End If%>
  </table>
<%
Case  "home"' home
%>
<%
End Select
'response.write time()-s1 
Set Session("cart") = dictCart
%>