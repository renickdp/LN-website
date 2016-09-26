<!--#INCLUDE FILE="inc/general.asp"-->
<% ' complete order --------------------------------------------------------------------------------
discount=session("discount")
if discount="" Then discount=0
if session("logonid")="" Then discount=0

sid=request.querystring("sid")
username=request.form("user")
shipcode=request.form("smt")
shipping=request.form("shipping")
customer=request.form("fullname")
If IsEmpty(customer) or customer="" then
	response.write "<center>Please enter a name."
	response.write "<form>"
	response.write "<input type='button' value='Retry' onclick=history.back()>"
	response.write "</form>"
	response.end
Else
customer=request.form("fullname")
End If
coaddress=request.form("coaddress")
If IsEmpty(coaddress) or coaddress="" then
	response.write "<center>Please enter a address ."
	response.write "<form>"
	response.write "<input type='button' value='Retry' onclick=history.back()>"
	response.write "</form>"
	response.end
Else
coaddress=request.form("coaddress")
End If

phone=request.form("phone")
If IsEmpty(phone) or phone="" then
	response.write "<center>Please enter a phone number ."
	response.write "<form>"
	response.write "<input type='button' value='Retry' onclick=history.back()>"
	response.write "</form>"
	response.end
Else
phone=request.form("phone")

End If

description=coaddress&"; "&phone&"; "& request.form("notes")

Function GetItemPrice(iItemID)
SQL = "SELECT* FROM Items WHERE code like '%" & iItemID & "%'"
rs.Open SQL, Conn
rs.MoveFirst
Dim aParameters
Do while not rs.eof
price=rs("Price")
aParameters=price-((price*discount)/100)
GetItemPrice = aParameters
rs.MoveNext
loop
rs.close
End Function

Function OrderNo()
SQL = "SELECT max(orderno) as neworderno FROM orders "
rs.Open SQL, Conn
rs.MoveFirst
Dim aParameters
Do while not rs.eof
neworderno = rs("neworderno")
if neworderno="" or neworderno=0 then neworderno=100000
OrderNo = neworderno+1
rs.MoveNext
loop
rs.close
End Function

Sub FinishCart()
Dim Key

Dim aParameters ' as Variant (Array)
orn=OrderNo()	
	For Each Key in dictCart
		'aParameters = GetItemParameters(Key)

SQLstmt = "INSERT INTO orders (sid,orderno,customer,adate,code,qty,price,shipping,shipcode,description,username)"
SQLstmt = SQLstmt & " VALUES ("
SQLstmt = SQLstmt & "'" & sid & "',"
SQLstmt = SQLstmt & "'" & orn & "',"
SQLstmt = SQLstmt & "'" & customer & "'," 
SQLstmt = SQLstmt & "'" & date() & "',"
SQLstmt = SQLstmt & "'" & key & "'," 
SQLstmt = SQLstmt & "'" & dictCart(Key) & "',"
SQLstmt = SQLstmt & "'" & GetItemPrice(Key) & "',"
SQLstmt = SQLstmt & "'" & shipping & "',"
SQLstmt = SQLstmt & "'" & shipcode & "',"
SQLstmt = SQLstmt & "'" & description & "',"
SQLstmt = SQLstmt & "'" & username & "'"
SQLstmt = SQLstmt & ")"
	
Response.write SQLstmt &"<br>"
Set RS = conn.execute(SQLstmt)
	
Next
response.cookies("qty") = 0
response.cookies("ord") = orn
'Session.Abandon
dictCart.RemoveAll
Session("qty")=0

mailserver="kiisu.aks.ee:25"

dim mailbody
saatja="orders@ln.com"
username="info@ln.com"


mailbody = " LN Shop Order # " & orn & Vbcrlf & vbcrlf
mailbody = mailbody & "http://LN.com/cart_v1/order.asp?orderno=" & orn & Vbcrlf & vbcrlf 
mailbody = mailbody & vbCrlf & vbcrlf &  "http://LN.com/cart_v1/" 
blabla = "subscription " & orn
Response.Write mailbody

set SMTP=Server.CreateObject("Jmail.SMTPMail")
SMTP.ServerAddress= mailserver
SMTP.Sender=BLAH
SMTP.AddRecipient username
SMTP.Subject=blabla
SMTP.Body=mailbody
SMTP.Execute
Session("login")=0

End Sub
 
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
iItemID = CStr(Request.QueryString("item"))
iItemCount = CInt(Request.QueryString("count"))

Select Case sAction
     	Case "finishcart"
FinishCart

End Select
'Set Session("cart") = dictCart

ord=request.cookies("ord")
Response.redirect("order.asp?orderno="& ord)
%>