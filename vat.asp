<% ' this page will switch language
go_back=request.servervariables("HTTP_REFERER")
vat=request.cookies("vat")
if vat="" Then vat=1
if vat=1  then 
vat=1.18
Else
vat=1
End if
response.cookies("vat")=vat
response.cookies("vat").expires = dateadd("d",7,date())
response.Redirect go_back
%>