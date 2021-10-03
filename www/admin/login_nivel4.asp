<!--#include file=".\config_site.inc"-->

<%
response.expires = 0


'session.timeout = Session("tempo")

If session("codigo") = ""  Then 	%>
<script language="JavaScript">
	   	msg = "           SESSÃO INCERRADA            \n";
		msg += "_____________________________\n";
		msg += "O seu id não confere, sessão incerrada, sai fora maluco.\n\n";
		alert(msg + "\n");
		top.location.href="logout.asp";
		
	   </script>
	   <%
	   else
	   
	Set user = Server.CreateObject("ADODB.Recordset")
	user.Open "SELECT * FROM usuarios WHERE id = " & session("codigo"), conn
	
	If user.Eof or user.Bof then 
	%>
		<script>
	   	msg = "           SESSÃO INCERRADA!!!            \n";
		msg += "_____________________________\n";
		msg += "O seu id não confere, sessão incerrada.\n\n";
		alert(msg + "\n");
		
		top.location.href="logout.asp";
		
	   </script>
	<%
	
	elseIf not user("id") = session("codigo") or user.EOF Then
	%>
		<script>
	   	msg = "           SESSÃO INCERRADA!!!            \n";
		msg += "_____________________________\n";
		msg += "O seu id não confere, sessão incerrada.\n\n";
		alert(msg + "\n");
		
		top.location.href="logout.asp";
		
	   </script>
	<%
	
	Elseif not user("tique_user") = Session("tique_user") Then
	
	%>
	<script>
	   	msg = "           SESSÃO INCERRADA            \n";
		msg += "_____________________________\n";
		msg += "O seu Tique não confere, sessão incerrada.\n\n";
		alert(msg + "\n");
		top.location.href='logout.asp';
	   </script>
	   <%
	   elseif session("Acesso") < 3 Then	
 session("erro") = "3"
response.redirect "login_erro.asp"

else
%>
<center>
	<table border="0" width="90%" cellspacing="0" cellpadding="0">
		<tr>
			
      <td width="34%"><img src="../imagens/logo_dvdex_mini.jpg" width="210" height="114"></td>
			<td width="66%">
			<center>
          <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#000000">Olá 
          <b><i><%=session("Nome")%></i></b>, você está logado com nível <b><%=session("acesso_escreve")%></b> <br>
          <font size="1">Tique de Sessão: <%=Session("tique_user")%></font></font> 
          <font face="Verdana, Arial, Helvetica, sans-serif"><br>
          <a href="usuarios_personaliza_admin.asp?usuarios=editar&codigo=<%=session("codigo")%>" ><font size="2">Clique 
          aqui para atualizar/editar seus dados</font></a></font> 
        </center>

</td>
		</tr>
	</table>
	</center>
<%


user.close
set user = nothing
Conn.Close
	set Conn = nothing
End If
End If
%>
