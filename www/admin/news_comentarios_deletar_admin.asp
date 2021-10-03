

<body onload="askData()" >

<%
id = Request.QueryString("ID")
id_noticia = Request.QueryString("id_noticia")
%>
<SCRIPT LANGUAGE="JavaScript">

function askData() {
//var inputedData =  prompt ("type something!", "" );
if (confirm("Tem certeza que deseja deletar ???")) {
//alert ("Ok, you did type "+inputedData+"!");
location.href('?del=ok&ID=<%=id%>&id_noticia=<%=id_noticia%>');
}
else {
location.href('news_comentarios_admin.asp?id=<%=id_noticia%>');
//alert ("No, you did not type "+inputedData+", did you?  OK.  Guess ya did.");
   }   
}
// End -->
</SCRIPT>
<!-- #include file="login_nivel3.asp" -->


<!--#include file=".\config_site.inc"-->



<%
if request("tabela") = "Dicas" Then
tabela = "Dicas"
else
tabela = "Noticias"
End If


if  Request.QueryString("del") = "ok" then
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open base_dados

 Sql="DELETE FROM "&tabela&"_comentario WHERE id=" & Request.QueryString("ID")
  Conn.Execute(Sql)
%>
  <script>alert('Dados delatados com sucesso!!!!!!  ');location.href='news_comentarios_admin.asp?id=<%=id_noticia%>&tabela=<%=tabela%>';</script>
   <%

  Conn.Close
  Set Conn = Nothing
 End If 
 
 'Response.redirect "locais_escalada_default_admin.asp" 
  
%>


