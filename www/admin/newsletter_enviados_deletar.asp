

<body onload="askData()" >

<%
id = Request.QueryString("ID")
%>
<SCRIPT LANGUAGE="JavaScript">

function askData() {
//var inputedData =  prompt ("type something!", "" );
if (confirm("Tem certeza que deseja deletar ???")) {
//alert ("Ok, you did type "+inputedData+"!");
location.href('?del=ok&ID=<%=id%>');
}
else {
location.href('newsletter_registros.asp?mostrar=<%=request.QueryString("mostrar")%>');
//alert ("No, you did not type "+inputedData+", did you?  OK.  Guess ya did.");
   }   
}
// End -->
</SCRIPT>
<!-- #include file="login_nivel3.asp" -->


<!--#include file=".\config_site.inc"-->




<%

if  Request.QueryString("del") = "ok" then
'Set Conn = Server.CreateObject("ADODB.Connection")
'Conn.Open base_dados


 Sql="DELETE FROM mail_list_send WHERE id=" & Request.QueryString("ID")
  Conn.Execute(Sql)
  
  
%>
  <script>alert('Dados delatados com sucesso!!!!!!  ');location.href='newsletter_registros.asp?p=False';</script>
   <%

  Conn.Close
  Set Conn = Nothing
 End If 
 
 'Response.redirect "locais_escalada_default_admin.asp" 
  
%>


