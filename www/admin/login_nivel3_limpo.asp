<!--#include file=".\config_site.inc"-->
<!--#include file="activeusers.asp"-->

<%
response.expires = 0
Set opa = Server.CreateObject("ADODB.Recordset")
	opa.Open "SELECT * FROM configuracoes_site", conn
	
If opa("tranca_site") = TRUE Then
	%>
		<script>
	   	msg = "             SISTEMA EM MANUTENÇÃO!!!            \n";
		msg += "____________________________________\n";
		msg += "O Sisitema esta em manutenção pelo adiministrador.\n\n Voltamos a funcionar dentro de alguns minutos aguarde!!! \n";
		alert(msg + "\n");
		
		top.location.href="tranca_site.asp";
		
	   </script>
	<%
	opa.close
	End If

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

user.close
set user = nothing
Conn.Close
	set Conn = nothing
End If
End If
%>
