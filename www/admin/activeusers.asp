<%
dim onlinedate
onlinedate = now()
OnlineUserIp = Request.ServerVariables("REMOTE_ADDR")
'Set timeout in minutes
strTimeout = 5
StrOnlineTimedout = dateadd("n",-strtimeout*60,onlinedate)

'Set Conn = Server.CreateObject("ADODB.Connection")
'Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("activeusers.mdb")
'Conn.Open

StrSql = "Select * From Active_Users Where iduser='"&session("codigo")&"'"
Set Rs1 = Conn.Execute (StrSql)

'Check if users ip exists in active users database
if rs1.eof then
'Add user to database because they dont exist
StrSql = "INSERT INTO Active_Users (Logon_time,User_ip,Last_Seen,Last_Page,nome,iduser) " &_
"VALUES ('" & onlinedate & "','" & OnlineUserIp & "','" & onlinedate & "','" &_
request.servervariables("path_info") & "','"&session("nome")&"','"&session("codigo")&"')"
Conn.Execute (StrSql)
else
'There Still active so lets update their info
StrSql = "UPDATE active_users SET Last_Seen='" & onlinedate & "', Last_Page='" &_
request.servervariables("path_info") & "', nome='"&session("nome")&"',iduser='"&session("codigo")&"' WHERE iduser='"&session("codigo")&"'"

'ser_ip='" & OnlineUserIp & "'"
Conn.execute (StrSql)
end if

'Remove Users Who Have Timed Out
StrSql = "DELETE FROM active_users WHERE Last_Seen < '" & StrOnlineTimedout & "'"
conn.execute (StrSql)

'Get Number of Active Users (Finally)
set rscount = server.createobject("ADODB.recordset")
strSql = "Select * from active_users"
rscount.open strSql,Conn,1,1
dim NumActiveUsers
NumActiveUsers = rscount.recordcount
%>