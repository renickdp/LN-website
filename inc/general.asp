<%
' connection to the db
' ------------------------------------------------
'Set Conn = Server.CreateObject("ADODB.Connection")
'Set Rs = Server.CreateObject("ADODB.Recordset")
    
'conn.Open "DBQ=" & Server.Mappath("db.mdb") & ";Driver={Microsoft Access Driver (*.mdb)};"

'DNS-less database connection
set rs=Server.CreateObject("adodb.Recordset")
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("inc/db.mdb") & ";"

'dsnname="DSN=intranet"
'set rs=Server.CreateObject("adodb.Recordset")
'connectme=dsnname

' ------------------------------------------------
'set conn = server.createobject("adodb.connection")
'dsnname="intranet"
' ------------------------------------------------
%>