<%
if request("tabela") = "Dicas" Then
tabela = "Dicas"
else
tabela = "Noticias"
End If
%>

<body onload="askData()" >

<%
id = Request.QueryString("ID")
%>
<SCRIPT LANGUAGE="JavaScript">

function askData() {
//var inputedData =  prompt ("type something!", "" );
if (confirm("Tem certeza que deseja deletar ???")) {
//alert ("Ok, you did type "+inputedData+"!");
location.href('?del=ok&ID=<%=id%>&tabela=<%=tabela%>');
}
else {
location.href('news_default_admin.asp?mostrar=<%=request.QueryString("mostrar")%>&tabela=<%=tabela%>');
//alert ("No, you did not type "+inputedData+", did you?  OK.  Guess ya did.");
   }   
}
// End -->
</SCRIPT>
<!-- #include file="login_nivel2.asp" -->


<!--#include file=".\config_site.inc"-->




<%

if  Request.QueryString("del") = "ok" then
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open base_dados


set Rs = server.createobject("ADODB.recordset")

Sqltext = "SELECT * FROM "&tabela&"_comentario WHERE id_noticia = '" & Request.QueryString("ID") & "' order by id DESC"
Rs.Open SqlText, base_dados

if Rs.EOF AND Rs.BOF =True Then
Else

Do Until Rs.EOF


	Sql2="DELETE FROM "&tabela&"_comentario WHERE id=" & Rs("ID")
  Conn.Execute(Sql2)
   Rs.MoveNext
   Loop
   
 End if
  Rs.Close
   Set Rs = Nothing
   
   
 Sql="DELETE FROM "&tabela&" WHERE id=" & Request.QueryString("ID")
  Conn.Execute(Sql)
  
  
%>
  <script>alert('Dados delatados com sucesso!!!!!!  ');location.href='news_default_admin.asp?p=False&tabela=<%=tabela%>';</script>
   <%

  Conn.Close
  Set Conn = Nothing
 End If 
 
 'Response.redirect "locais_escalada_default_admin.asp" 
  
%>


