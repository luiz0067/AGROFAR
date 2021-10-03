<!-- #include file="login_nivel3_limpo.asp" -->
<!--#include file=".\config_site.inc"-->
<%
if request("m")="del"Then


 	pe = "SELECT * FROM menu_item"' WHERE id_menu = "&request("id_menu")
	Set rs = Conn.Execute(pe)
  if Rs.EOF AND Rs.BOF =True Then
		Else
			Do Until Rs.EOF
				if rs("id_menu") = request("id_menu") Then
					Sql2="DELETE FROM menu_pgs WHERE id=" & Rs("ID")
  					Conn.Execute(Sql2)
				End If
   			Rs.MoveNext
   		Loop
	End if
  Rs.Close
   Set Rs = Nothing

 Sql="DELETE FROM menu WHERE id=" & Request("id_menu")
  Conn.Execute(Sql)
  
  Sql2="DELETE FROM menu_item WHERE id_menu=" & Request("id_menu")
  Conn.Execute(Sql2)
  
    Sql3="DELETE FROM menu_pgs WHERE id_menu=" & Request("id_menu")
  Conn.Execute(Sql3)
 
%>
Tudo relacionado a este menu foi  [<%=request("menu")%>] DELETADO com Sucesso!!!<br>
<br>
<script>
window.opener.location.href="menus_default.asp";
self.close();
//alter("Menu Cadastrado com Sucesso!!");location.href="";</script>
<%	End If	%>
