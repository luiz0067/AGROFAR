

<!-- #include file="login_nivel2.asp" -->
<%


Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open base_dados

 Sql="DELETE FROM chamada WHERE id=" & Request("id")
  Conn.Execute(Sql)
response.Redirect("chamada_default_admin.asp?msg=3")
 
%>

