
<!--#include file=".\config_site.inc" -->

<!-- Desenvolvido por Midiamix-->
<title>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;..:: C R M - system ::..  &nbsp;&nbsp;&nbsp;(( <%=site_titulo%> ))&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr > 
    <td height="25"  background="../imagens/forum_images/table_bg2.gif" bgcolor="#EEEEEE">&nbsp;</td>
    <td width="35%" valign="middle"  background="../imagens/forum_images/table_bg2.gif" bgcolor="#EEEEEE"> 
      <div align="right"><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Seja 
        bem vindo(a) ao sistema, <%=now()%>&nbsp;&nbsp;&nbsp;&nbsp;</strong></font></div></td>
  </tr>
  <tr  > 
    <td height="8" colspan="2" background="../imagens/outros_icos/spacer-news.gif"></td>
  </tr>
  <tr > 
    <td height="113" colspan="2" bgcolor="cor_fundo_site"> <table width="100%"  height="100%"border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="31%" height="101" bgcolor="#ffffff"><div align="center"><img src="../<%=logo_empresa%>" >&nbsp;</div></td>
          <td width="1%" background="../imagens/outros_icos/sm-tree1.gif"></td>
          <td width="68%" bgcolor="<%=cor_fundo_site%>">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr  > 
    <td  height="8" colspan="2" background="../imagens/outros_icos/spacer-news.gif"></td>
  </tr>
  <tr > 
    <td colspan="2"  height="8"><br><br>
<% Response.Buffer = True %>
<%
	Login = Trim(Request.Form("Login"))
	Senha = Trim(Request.Form("Senha"))
	Verifica = 0

	Set Connexao = Server.CreateObject("ADODB.Connection")
	Connexao.Open base_dados
	
	SQLString = "SELECT * FROM usuarios WHERE login='" & Login & "'"

	Set Reg = Connexao.Execute(SQLString)
	If Reg.EOF Then
		Verifica = 1
	End If
%>

	<TITLE>:: Autenticação de Usuário ::</TITLE>

</HEAD>

<BODY BGCOLOR=#FFFFFF>
<% If Verifica = 1 Then %>
<script>alert('Login incorreto, por favor tente outra vez  ');location.href='default.asp';</script>

<% Else %>
<% If Reg("Senha") <> Senha Then %>
<script>alert('Senha incorreta, por favor tente outra vez  ');location.href='default.asp';</script>

<% Else
'Session("tempo") = "50"
'tempo = Session("tempo")
		'gerando um tique para a sessão do usuário 
		randomize timer	
	tique_user = Left(Request("login"),3) & (9876989856 * CInt((RND * 32000) + 100 )) & right(Request("login"),2)
	
	Sql = "UPDATE usuarios SET tique_user='" & tique_user & "' WHERE id=" & Reg("id")
  	Conn.Execute(Sql)
	
	session("tique_user") = tique_user
	session("Nome")   = Reg("Nome")
	session("Acesso") = Reg("Acesso")
	session("codigo") = Reg("id")
	session("foto") = Reg("foto")
	'session("alert_msg") = Reg("alert_msg")
	nome_user_log = session("Nome") 

			if  session("acesso") = "1" then
			session("acesso_escreve") = "Básico"
			end if
			if  session("acesso") = "2" then
			session("acesso_escreve") = "Limitado"
			end if
			if session("acesso") = "3" then
			session("acesso_escreve") = "Administrador"
			end if

	%>

	<meta http-equiv="refresh" content="3;URL=menu_meio.asp">
      <div align="center">
        <style>
<!--
.clsNormal    { font-family: Verdana, Arial, Helvetica, MS Sans Serif, sans-serif; font-size:12px; padding: .2em;border-color: #000000; }
.clsDynaText    { font-family: Verdana, Arial, Helvetica, MS Sans Serif, sans-serif; padding: .2em; font-size:12px; color:#ffffff; border-color: #000000;background-color:#6495ed; }
-->
</style>
        <SCRIPT LANGUAGE="JavaScript">
<!--
    var sPlainText 
    var iTick = 35
    var j = 0
    var sOffID
    
    window.onload = doOnload
    
    function doOnload() {
        idDynaText.innerHTML = replaceChars(idDynaText.innerText," ","&nbsp;")
        sPlainText = idDynaText.innerText
        getString()
    }
    
    function getString() {
        sDynaText = ""
        for (i=0;i<sPlainText.length;i++){
            sDynaText += "<SPAN ID=idDynaChar" + i + " CLASS=clsNormal>" + sPlainText.substring(i,i+1) + "</SPAN>"
        }
        idDynaText.innerHTML = sDynaText
        window.setTimeout("doDynaText()",iTick)
    }

    function doDynaText() {
        if (null != sOffID) {
        }
        sOnID = "idDynaChar" + (j)
        document.all(sOnID).className = "clsDynaText"
        sOffID = sOnID
        j++
        if (j < sPlainText.length    ) {
            window.setTimeout("doDynaText()",iTick)
        }
        else {
            if (null != sOffID) {
            }
        }
    }

    function replaceChars(sString,sOld,sNew) {
        for (i = 0; i < sString.length; i++) {
            if (sString.substring(i, i + sOld.length) == sOld) {
                sString = sString.substring(0, i) + sNew + sString.substring(i + sOld.length, sString.length)
            }
        }
        return sString
    }
-->
</SCRIPT>
<!--#include file="activeusers.asp"-->
        <center>
          <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Usuário 
          autenticado com Sucesso!!!<br>
          Aguarde. Carregando configurações do sistema...<br>
          Caso o sistema não redirecione automaticamente <a href="menu_meio.asp">clique 
          aqui</a>. </strong></font><br>
          <br>
		  <br> <span id="idDynaText"  > 
          <table width="71%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td height="27">&nbsp;<font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Usuário 
                autenticado com Sucesso!!! Aguarde, carregando configurações....</strong></font></td>
            </tr>
          </table>
          </span> 
        </center>
		<br><br>
		<br>
		<br>
        <%
	'response.redirect "menu_frames.asp"

End If %>
        <% End If

	Reg.Close
	Connexao.Close
	Set Reg = Nothing
	Set Connexao = Nothing

%>
      </div></td>
  </tr>
  <tr > 
    <td  height="8" background="../imagens/outros_icos/spacer-news.gif"  colspan="2"></td>
  </tr>
  <tr > 
    <td colspan="2"  height="8"><div align="right"><font color="#666666" size="1" face="Arial, Helvetica, sans-serif">Desenvolvido 
        por Midiamix&nbsp;&nbsp;&nbsp;&nbsp;<br>
        &copy; 2006 - Todos os direitos preservados&nbsp;&nbsp;&nbsp;&nbsp;</font></div></td>
  </tr>
</table>

