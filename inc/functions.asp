<%user = request.cookies("passes")%>
<%fullname = request.cookies("passes2")%>
<%If request.cookies("passes") = "" then Session("logonid")="" %>
<%if session("logonid")="" then response.cookies("passes")=""%>
<%'session.Timeout=1%>
<% p_url=request.servervariables("SCRIPT_NAME")
url=request.servervariables("SCRIPT_NAME")
sss=request.servervariables("HTTPS")
'if sss="off" Then response.redirect("https://ln.com/cart_v1/")
' ------------------------------------------------
homepage="default.asp"
mainpage="index.asp"
offerspage="offer.asp"
s1=time()
vat=request.cookies("vat")
dc=request.cookies("dc")

' ------------------------------------------------

l=Session("login")
qty=Session("qty") 
sid=(Session.SessionID)
' ------------------------------------------------
Sub CountOfferCart()
	For Each Key in dictOffer
		q= q + dictOffer(Key) 
	 Next
	session("offerqty")= Cint(q)
End Sub
' ------------------------------------------------

Sub CountCart()
'if qty="" Then qty=0
	For Each Key in dictCart
		q= q + dictCart(Key) 
	 Next
	session("qty")= Cint(q)
End Sub

if request.cookies("qty")="" Then response.cookies("qty")=0
' ------------------------------------------------
if request.cookies("vat")="" Then response.cookies("vat")=1
response.cookies("vat").expires=dateadd("d",7,date())

' ------------------------------------------------
discount=session("discount")
if discount="" Then discount=0
if session("logonid")="" Then discount=0
' ------------------------------------------------
' ------------------------------------------------
if session("admin")="" Then session("admin")=0
' ------------------------------------------------
if request.cookies("dc")="" Then response.cookies("dc")=1
response.cookies("dc").expires=dateadd("d",7,date())

' ------------------------------------------------
sURL = Request.ServerVariables("SCRIPT_NAME")
if Request.ServerVariables("QUERY_STRING") <> "" Then
sURL = sURL & "?" & Request.ServerVariables("QUERY_STRING")
End if
' ------------------------------------------------
if request.form("m")="" Then
m=request.querystring("m")
Else
m=request.form("m")
End If
if m="" Then m=1
' ------------------------------------------------

' 
function LinkURLs(ByRef asContent)
    	Dim loRegExp	' Regular Expression Object (Requires vbScript 5.0 and above)
    	' if no content was received, Exit the function
    	if asContent = "" Then Exit function
    	' Create Regular Expression object
    	Set loRegExp = New RegExp
    	' Keep finding links after the first one.				
    	loRegExp.Global = True
    	' Ignore upper/lower Case
    	loRegExp.IgnoreCase = True
    	' Look For URLs
    	loRegExp.Pattern = "((ht|f)tps?://\S+[/]?[^\.])([\.]?.*)"
    	' Link URLs
    	LinkURLs = loRegExp.Replace(asContent, "<A href=""$1"" target='_blank'>$1</A>$3")
    	'LinkURLs = loRegExp.Replace(asContent, "<A href=""$1""><img border='0' src='../img/www_ico.gif' alt='$1'></A>$3")
    	' Look For email addresses
    	loRegExp.Pattern = "(\S+@\S+.\.\S\S\S?)"
    	' Link email addresses
    	LinkURLs = loRegExp.Replace(LinkURLs, "<A href=""mailto:$1"">$1</A>")
    	    	
    	' Release regular expression object
    	Set oRegExp = Nothing
    	
    End function
%>