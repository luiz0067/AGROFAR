<!-- #include file="login_nivel3_limpo.asp" -->
<%
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open base_dados
'StrSql = "DELETE FROM active_users WHERE Last_Seen < '" & StrOnlineTimedout & "'"
StrSql = "DELETE FROM active_users WHERE id= " & request("id")
conn.execute (StrSql)

'Get Number of Active Users (Finally)
response.Redirect("usuarios_default_admin.asp")
%>