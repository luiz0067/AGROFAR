
<!-- #include file="login_nivel3.asp" -->
<!--#include file=".\config_site.inc"-->
<%

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open base_dados

if  Request("acao") = "Deletar E-mails Selecionados" then

Faz = " Delete From "
Else
Faz = " Update "
End If

if  Request("acao") = "Setar E-mails Offline" then
setar = " SET online = FALSE "
escreve = "Offline"
End If

if  Request("acao") = "Setar E-mails Online" then
setar = " SET online = TRUE "
escreve = "Online"
End If

'Response.Write setar
'Response.Write faz
    For Each I in Request.Form("id")
	Response.Write I & "<br>"
        Sql= Faz & "mail_list" & setar & " WHERE ID= " & I
		Conn.Execute(Sql) 
   Next
   Conn.Close
   Set Conn = Nothing

if  escreve = "" then
%>
		<script>
		msg = "           DADOS DELETADOS COM SUCESSO!!!                    \n\n";
		msg += "_____________________________________________________________\n\n";
		msg += "O(s) E-mail(s) foram deletados do banco de dados com sucesso.\n";
		alert(msg + "\n\n");
		location.href='newsletter_delete.asp?mostrar=<%=Request("mostrar")%>&PageNumber=<%=request("pg")%>';
		</script>
<%	else	%>
		<script>
		msg = "           DADOS ATUALIZADOS COMO <%=escreve%> SUCESSO!!!                    \n\n";
		msg += "_____________________________________________________________\n\n";
		msg += "O(s) E-mail(s) foram setados como <%=escreve%> no banco de dados.\n";
		alert(msg + "\n\n");
		location.href='newsletter_delete.asp?mostrar=<%=Request("mostrar")%>&PageNumber=<%=request("pg")%>';
		</script>
<%	End If	%>


