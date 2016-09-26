<%if Session("logonid") = "" and l<4 Then  %>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
    <form method="POST" action="login.asp">
  <INPUT type="hidden" name="dologin" value="yes">
<%  Dim sURL
    	sURL = Request.ServerVariables("SCRIPT_NAME")
    	if Request.ServerVariables("QUERY_STRING") <> "" Then
    	sURL = sURL & "?" & Request.ServerVariables("QUERY_STRING")
   	End if%>
  <INPUT type="hidden" name="siteid" value="<%=surl%>">
  <tr>
    <td>
    <p align="right">Kasutajanimi :&nbsp; </td>
    <td>
         <p>
         <input type="text" name="id" size="10" class="boxrequired" value="<%=request.cookies("passes")%>"></p>
     </td>
  </tr>
  <tr>
    <td>
    <p align="right">Salasõna :&nbsp; </td>
    <td>
      <input type="password" name="pwd" size="10" class="boxrequired"></td>
  </tr>
  <tr>
       <td align="center" colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
       <input type="image" src="img/btn_ok.gif" value="Submit" name="I1"></font></td>
  </tr>
  </form>
</table>
<%Else%>
    <%If l>3 Then 
    Response.write "konto lukustatud "
    Else
    Response.write fullname 
    End if %>
    <%End If%>