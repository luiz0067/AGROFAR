<link href="../css/css_admin.css" type="text/css" rel="stylesheet" />
<!--#include file=".\config_site.inc"-->
<!--#include file="activeusers.asp"-->
<style>
     	a {  text-decoration: none; color: #FF4646}
		a:hover {  text-decoration: underline}
 </style>
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
	   elseif session("Acesso") < 1 Then	
 session("erro") = "2"
response.redirect "login_erro.asp"

else
%>


<!-- Desenvolvido por Midiamix -->
<title>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;..:: C R M - system ::..  &nbsp;&nbsp;&nbsp;(( <%=site_titulo%> ))&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr > 
    <td height="25"  background="imagens/forum_images/table_bg2.gif" bgcolor="#EEEEEE"><font face="Arial, Helvetica, sans-serif" size="2" color="#FF0000"> 
      &nbsp;&nbsp;&nbsp;&nbsp;Olá <b><i><%=session("Nome")%></i></b>, você está 
      logado com nível <b><%=session("acesso_escreve")%></b> <font size="1">| 
      Cod.Sessão: <%=Session("tique_user")%> |</font></font></td>
    <td width="20%" valign="middle"  background="imagens/forum_images/table_bg2.gif" bgcolor="#EEEEEE"> 
      <div align="right"><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="logout.asp"><img src="imagens/bt_sair.gif" alt="Clique aqui para Sair do Sistema" width="58" height="20" border="0"  align="absmiddle"></a>&nbsp;&nbsp;&nbsp;<%=now()%>&nbsp;&nbsp;&nbsp;&nbsp;</strong></font></div></td>
  </tr>
  <tr  > 
    <td height="8" colspan="2" background="imagens/outros_icos/spacer-news.gif"></td>
  </tr>
  <tr > 
    <td height="113" colspan="2" bgcolor="cor_fundo_site"> <table width="100%"  height="100%"border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="31%" height="101" bgcolor="#ffffff"><div align="center"><img src="../<%=logo_empresa%>" ></div></td>
          <td width="1%" background="imagens/outros_icos/sm-tree1.gif"></td>
          <td width="68%" align="center" valign="middle" bgcolor="<%=cor_fundo_site%>"><!--#include file="topo_botoes.asp"--> 
          </td>
        </tr>
      </table></td>
  </tr>
  <tr  > 
    <td  height="8" colspan="2" background="imagens/outros_icos/spacer-news.gif"></td>
  </tr>
  <tr  >
    <td   colspan="2" ><hr color="<%=cor_links%>"></td>
  </tr>
</table>

<br>
<%
 ' Alerta sobre chegada de e-mails
'alertMsg
user.close
set user = nothing
'Conn.Close
'	set Conn = nothing
End If
End If
%>
