
<!-- #include file=".\config_site.inc" -->
<!-- #include file="login_nivel2.asp" -->
<%
if request("tabela") = "Dicas" Then
tabela = "Dicas"
else
tabela = "Noticias"
End If
r1 = "|_aspas_|"
r2 = "'"
msg = Request("Body")
msg = replace(msg,"|_aspas_|","'")
msg = replace(msg,"'","''") 
'msg = msg2




Set Conn = Server.CreateObject("ADODB.Connection")
   Conn.Open base_dados
   data_hoje = Day(Now) & ("/") & month(Now) & ("/") & Year(Now)
   
   if Request("Submit") = "Cadastrar" Then
   
    Sql="INSERT INTO "&tabela&" (titulo,autor,mailautor,texto,contagem,recomendado,cabeca,data,online)"
	 sql = sql + "VALUES('" & Request("titulo") & "','" & Request("autor") & "','" & Request("mailautor") & "',"
	 sql = sql + "'" & msg & "',0,0,'" & Request("cabeca") & "','" & data_hoje & "',TRUE)"
   Conn.Execute(Sql)
   %>
   <script>
	   	msg = "           <%=tabela%> CADASTRADA COM TOTAL SUCESSO!!!              \n";
		msg += "________________________________________________________________\n\n\n";
		msg += "A sua <%=tabela%>  foi cadastrada no nosso banco de dados com Sucesso.\n\n";
		msg += "Você está sendo redirecionado para a página inicial de <%=tabela%> .\n\n";
		alert(msg + "\n\n");
		location.href="news_default_admin.asp?tabela=<%=tabela%>";
   </script>
   <%
   Else
   Sql="Update "&tabela&" SET titulo = '" & Request("titulo") & "', autor = '" & Request("autor") & "',"
   sql = sql + "mailautor = '" & Request("mailautor") & "', texto = '" & msg & "',"
   sql = sql + "contagem = '" & Request("contagem") & "', recomendado = '" & Request("recomendados") & "',"
   sql = sql + "cabeca = '" & Request("cabeca") & "',data = '" & Request("data") & "', online=TRUE"
   sql = sql + " WHERE ID=" & request("id")
		Conn.Execute(Sql)
		%>
		  <script>
	   	msg = "          <%=tabela%>  ATUALIZADA COM TOTAL SUCESSO!!!              \n";
		msg += "________________________________________________________________\n\n\n";
		msg += "A <%=tabela%>  <%=Request("titulo")%> \n";
		msg += "foi Atualizada no nosso banco de dados com Sucesso.\n\n";
		msg += "Você está sendo redirecionado para a página inicial de <%=tabela%> .\n\n";
		alert(msg + "\n\n");
		location.href="news_default_admin.asp?tabela=<%=tabela%> ";
   </script>
		<%
  ENd If 
  
%>