<!-- #include file="login_nivel3.asp" -->

<%
if request("tabela") = "Dicas" Then
tabela = "Dicas"
else
tabela = "Noticias"
End If

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open base_dados
   
if  Request.QueryString("ativar") = "sim" then
    For Each I in Request.Form
     a = Request.Form(I)
	 If Request.Form(I) <> "Atualizar" Then
	    Sql="Update produtos SET online=TRUE WHERE ID=" & a
		Conn.Execute(Sql)
     End IF			 
   Next
   Conn.Close
   Set Conn = Nothing
   Response.redirect "produtos_default.asp?mostrar=off-line&tabela="&tabela
else
 
 For Each I in Request.Form
     a = Request.Form(I)
	 If Request.Form(I) <> "Atualizar" Then
	    Sql="Update produtos SET online=FALSE WHERE ID=" & a
		Conn.Execute(Sql)
     End IF			 
   Next
   Conn.Close
   Set Conn = Nothing
   Response.redirect "produtos_default.asp?tabela="&tabela
   
end if
   

%>