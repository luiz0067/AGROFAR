<!-- #include file="login_nivel3.asp" -->

<!--#include file=".\config_site.inc"-->

<% IF REQUEST("ac")= "cadastrar" then
Sql="INSERT INTO menu_pgs (titulo,subtitulo,conteudo,id_menu) VALUES('" & Request("titulo") & "','" & Request("subtitulo") & "','" & Request("body") & "',"&request("id_menu")&")"
  Conn.Execute(Sql)
  Sql="Update menu SET pg_ref= TRUE WHERE id ="& Request("id_menu")
   Conn.Execute(Sql)  
%>
P�gina  do Menu [<%=request("menu")%>] Cadastrada com Sucesso!!!<br>
<br>
<a href="menus_default.asp">Voltar para � P�gina de Menus</a> 
<% End If	%>

<% IF REQUEST("ac")= "editar" then
Sql="Update menu_pgs SET titulo='" & Request("titulo") & "',subtitulo='" & Request("subtitulo") & "',conteudo='" & Request("body") & "' WHERE id_menu="&request("id_menu")
  Conn.Execute(Sql)
  'Sql="Update menu SET pg_ref= TRUE WHERE id ="& Request("id_menu")
   'Conn.Execute(Sql)  
%>
P�gina  do Menu [<%=request("menu")%>] Atualizada com Sucesso!!!<br>
<br>
<a href="menus_default.asp">Voltar para � P�gina de Menus</a> 
<% End If	%>

<!-- #include file="rodape.asp" -->
