<!--#include file=".\config_site.inc"-->

<% IF REQUEST("ac")= "cadastrar" then
Sql="INSERT INTO produtos (tit_menu,titulo,subtitulo,conteudo,online) VALUES('" & Request("tit_menu") & "',"
Sql=Sql+" '" & Request("titulo") & "','" & Request("subtitulo") & "','" & Request("body") & "',TRUE)"
  Conn.Execute(Sql)
  
%>
P�gina de Produto Cadastrado com Sucesso!!!<br>
<br>
<a href="produtos_default.asp">Voltar para � P�gina de Produtos</a> 
<% End If	%>

<% IF REQUEST("ac")= "editar" then
Sql="Update produtos SET titulo='" & Request("titulo") & "',subtitulo='" & Request("subtitulo") & "',"
Sql=Sql+"conteudo='" & Request("body") & "',tit_menu='"&request("tit_menu")&"' WHERE id ="&request("id")
  Conn.Execute(Sql)
%>
P�gina  de Produtos Atualizada com Sucesso!!!<br>
<br>
<a href="produtos_default.asp">Voltar para � P�gina de Menus</a> 
<% End If	%>

