<!--#include file=".\config_site.inc"-->
<%
if request("it")="del"Then

 Sql="DELETE FROM menu_item WHERE id=" & Request("ID")
  Conn.Execute(Sql)
  
  Sql="DELETE FROM menu_pgs WHERE id_menu_it=" & Request("ID")
  Conn.Execute(Sql)
%>
Ítem do Menu [<%=request("menu")%>] DELETADO com Sucesso!!!<br>
<br>
<script>
window.opener.location.href="menus_default.asp";
self.close();
</script>
<%	End If	%>
