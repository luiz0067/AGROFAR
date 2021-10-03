 <!--#include file ="config_site.inc"-->
 
 <!-- Desenvolvido por Midiamix-->
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
      <div align="right"><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=now()%>&nbsp;&nbsp;&nbsp;&nbsp;</strong></font></div></td>
  </tr>
  <tr  > 
    <td height="8" colspan="2" background="imagens/outros_icos/spacer-news.gif"></td>
  </tr>
  <tr > 
    <td height="113" colspan="2" bgcolor="cor_fundo_site"> <table width="100%"  height="100%"border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="31%" height="101" bgcolor="#ffffff"><img src="<%=logo_empresa%>" ></td>
          <td width="1%" background="imagens/outros_icos/sm-tree1.gif"></td>
          <td width="68%" align="center" valign="middle" bgcolor="<%=cor_fundo_site%>"> 
            <!--#include file="includes/topo_botoes.asp"--></td>
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
<table width="50%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td><div align="center"><img src="imagens/chave.gif" width="75" height="75" align="middle"><font color="#FF0000" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Aacute;rea 
        restrita &agrave; usu&aacute;rios de N&iacute;vel <%=Session("erro")%></strong></font></div></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
</table>
<p align="center"> 
  <% if Session("erro") = "3" then %>
<p align="center"><font face="Arial" size="3"><b>
   Desculpe<font color="#FF0000"> <i><%=session("Nome")%></i></font>, mas está área é restrita a usuários de nível <font color="#FF0000">3</font>, você possui nível<font color="#FF0000">
    <%=session("Acesso")%></font>.<br>
Solicite ao webmaster autorização para ter acesso a está área.</b></font><br>
   
   
	<% end if %>
	
	
	<% if Session("erro") = "2" then %>

<p align="center"><font face="Arial" size="3"><b>
   Desculpe <font color="#FF0000"><b><i><%=session("Nome")%></i></b></font>, mas está área é restrita a usuários de nível <font color="#FF0000">2</font>, você possui nível<font color="#FF0000">
    <%=session("Acesso")%></font>.<br>
Solicite ao webmaster autorização para ter acesso a está área.</b></font><br>
   
   
	<% end if %>

	<% if Session("erro") = "1" then %>

<p align="center"><font face="Arial" size="3"><b>
   Desculpe <font color="#FF0000"><b><i><%=session("Nome")%></i></b></font>, mas está área é restrita a usuários de nível <font color="#FF0000">1</font>, você possui nível<font color="#FF0000">
    <%=session("Acesso")%></font>.<br>
Solicite ao webmaster autorização para ter acesso a está área.</b></font><br>
   
   
	<% end if %>
	
	<% if Session("erro") = "4" then %>

<p align="center"><font face="Arial" size="3"><b>
   Desculpe <font color="#FF0000"><b><i><%=session("Nome")%></i></b></font>, mas você ficou sem utilizar o sistema por mais de 5 minutos.<br>
<a href="logout.asp" target="_top">Clique aqui para identificar-se novamente</a>.</b></font><br>
   
   
	<% end if %>
	
	<% Session("erro") = "" %>
	
	<%
	conn.close
	set conn = Nothing
	%>
	 <!--#include file="includes/rodape.asp"-->