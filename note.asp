<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1257">
<title>Info lisamine</title>
</head>

<body>
<% ' add note --------------------------------------------------------------------------------
set rs=Server.CreateObject("adodb.Recordset")
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("inc/db.mdb") & ";"

Function EditNote(id)
SQL = "SELECT* FROM orders WHERE id = '%" & id & "%'"
rs.Open SQL, Conn
rs.MoveFirst
Do while not rs.eof
note = rs("note")
rs.MoveNext
loop
rs.close
End Function

Function AddNote(id)

End Function
 

%>
</body>

</html>