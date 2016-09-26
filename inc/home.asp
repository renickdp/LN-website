<%%>
<table border="0" cellpadding="0" cellspacing="0" width="770" class="maintable">
  <tr>
    <th width="100%" bgcolor="#add8e6">Welcome</th>
  </tr>
  <tr>
    <td width="100%" valign="top">
    <table border="0" cellpadding="0" class="indextable" width="770" background="../img/bg_cntr.gif" cellspacing="10" height="365">
      <tr>
        <td width="200" height="2" bgcolor="#FF9900" colspan="3"></td>
      </tr>
      <tr>
        <td width="200" valign="top" height="308" bgcolor="#F8F8F8">
        <img alt border="0" src="img/logo.gif"><br>
        &nbsp;<p>In Our Store :</p>
        <ul>
          <li>Science</li>
          <li>Ebgineering</li>
          <li>Medicine</li>
        </ul>
        <p>&nbsp;</td>
        <td width="370" valign="top">
        <font size="1"><br>
        </font><font color="#003399"><b>Introduction</b></font><p>
        <font color="#003399">This is a Rental Book store.</font></p>
        <p><b><font color="#003399">CURRENTLY CODED FEATURES</font></b></p>
        <ul>
         
          <li><font color="#003399">Products search by SKU or by groups</font></li>
          <li><font color="#003399">Personalized orders status tracing</font></li>
          <li><font color="#003399">Password reminder</font></li>
        </ul>
        <p><font color="#003399">Please 
        register your account<br>
&nbsp;<br>
        userid: admin@admin.com <br>
        password :admin</font></td>
        <td width="200" valign="top" bgcolor="#FFFFCC">
        <table border="0" cellpadding="0" cellspacing="0" class="maintable" width="100%" style="border-collapse: collapse" bordercolor="#111111">
          <tr>
            <th width="100%">This week favorite books</th>
          </tr>
          <tr>
            <td width="100%"><font size="1">&nbsp;</font><!--#INCLUDE FILE="specials.asp"--></td>
          </tr>
<%if Session("logonid") = ""  and l< 4 Then %>  
        <tr>
              <th width="100%">Log in </th>
          </tr>
          <tr>
            <td width="100%">
            <table border="0" class="maintable" width="100%" style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0">
            <form method="POST" name="loginformsmall" action="login.asp">
  <INPUT type="hidden" name="dologin" value="yes">
<%    	sURL = Request.ServerVariables("SCRIPT_NAME")
    	if Request.ServerVariables("QUERY_STRING") <> "" Then
    	sURL = sURL & "?" & Request.ServerVariables("QUERY_STRING")
   	End if%>
  <INPUT type="hidden" name="siteid" value="<%=surl%>">
  <tr>
    <td align="right" width="50%">
    User id:&nbsp; </td>
    <td width="50%">
         <p>
         <input type="text" name="id" size="10" class="boxrequired" value="user"></p>
     </td>
  </tr>
             <td align="right" width="50%"> password :&nbsp; </td>
    <td width="50%">
      <input type="password" name="pwd" size="10" class="boxrequired"></td>
  </tr>
  <tr>
       <td align="center" colspan="2">&nbsp;</td>
  </tr>
  <tr>
       <td align="center" colspan="2"> 
       <input type="image" src="img/btn_ok.gif" value="Submit" name="I1" border="0"></td>
  </tr>
  <tr>
    <td align="center" colspan="2" width="50%">&nbsp;</td>
          </form>
          </table>
          </tr>
          <tr>
            <th width="100%">Register</th>
          </tr>
          <tr>
            <td width="100%"><font size="1">Resgister and get 3 % discount for 
            each order </font><ul>
              <li><a href="index.asp?action=register"><font size="1">Register</font></a></li>
            </ul>
            <p>&nbsp;</td>
          </tr>
          <%End If%>
          <tr>
            <td width="100%">&nbsp;</td>
          </tr>
          </table>
        </td>
      </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td width="100%" height="2" bgcolor="#FF9900"></td>
  </tr>
</table>