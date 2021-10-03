<!--#include file="login_nivel3.asp"-->
<div align="center">
  <!--#include file="../newsletter_mail_func.asp"-->
  <%
email = Request.form("email")
estado = Request.QueryString("estado")
cidade = Request.form("cidade")
nome = Request.form("nome")
telefone  = Request.form("telefone")
titulo= Request.form("titulo")
comentario = Request.form("comentario")
%>
  <input name="button"  type="button" 
  style="cursor:hand;font-size:9pt;color: #CC0000; "   
  onClick="window.open('conta_mail_admin.asp','WinExa','titlebar=no,toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=0,width=320,height=150');"  
  value="Cadastrar Setores ">
  <input name="button2"  type="button" 
  style="cursor:hand;font-size:9pt;color: #CC0000; "   
  onClick="window.open('conta_tit_admin.asp','WinExa','titlebar=no,toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=0,width=320,height=150');"  
  value="Cadastrar Títulos ">
  <input name="button22"  type="button" 
  style="cursor:hand;font-size:9pt;color: #CC0000; "   
  onClick="location.href='contato_form.asp';"  
  value="Cadastrar Coneúdo de Página ">
  <br>
  Use os botões acima para cadastrar e montar a página de contato do site. Abaixo 
  você pode visualizar como a mesma esta ficando. <br>
  Vale lembrar que está página é padrão do sistema no link de menu "Contato". 
  <br>
  <br>
</div>
<table width="620" border="0" background="../imagens/img_fundo/bg_box2.jpg" align="center">
  <tr> 
    <td>&nbsp;</td>
    <td  valign="top" width="177"> <table width="177" border="0" height=37 cellspacing="0" cellpadding="1">
        <tr> 
          <td valign="top" background="../imagens/img_fundo/barratit.jpg"> <SPAN 
                        class=titulopagina><img src="../imagens/img_fundo/transp.gif" width="8" height="7">Contato</SPAN><br> 
            <span  class="subsubtit"><img src="../imagens/img_fundo/transp.gif" width="8" height="7">E-mail, 
            Endereços</span></td>
        </tr>
      </table></td>
    <td width="415" rowspan="2" valign="top"> 
      <%
						Sqltext2 = "SELECT * FROM configuracoes_site where id_geral = 1"
						Set Rs2 = Conn.Execute(Sqltext2)
						%>
      <div align="justify"><span class="verdebarra"><%=Rs2("contato_txt")%></span> 
        <%	
						Rs2.Close
						Set Rs2=nothing
						%>
      </div></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td  valign="top"><img src="../imagens/cara_mail3.gif" width="272" height="144"></td>
  </tr>
  <tr> 
    <td width="14">&nbsp;</td>
    <td colspan="2" valign="middle"> <form action="fale_obr.htm" enctype="application/x-www-form-urlencoded" method="post">
        <!-- -->
        <div align="center"> 
          <table border="0" cellspacing="8" cellpadding="0" width="376" height="208" class="bgformcontato" align="center"
		  >
            <tr valign="bottom"> 
              <td> <span class="chamadas">Escolha o setor<span class="txtmenordourado">*</span><br>
                <img src="../imagens/img_fundo/transp.gif" width="8" height="7"><br>
                </span> <table border="0" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td> <select name="para" class="txtmenor" style="width:160px;"  >
                        <option selected>Escolha...</option>
                        <%
						Sqltext2 = "SELECT * FROM contato_mails ORDER BY descri ASC"
						Set Rs2 = Conn.Execute(Sqltext2)
						If Rs2.EOF AND Rs2.BOF = TRUE Then
						Else
						Do While NOT Rs2.EOF
						%>
                        <option value="<%=Rs2("mail")%>"><%=Rs2("descri")%></option>
                        <%
						Rs2.MoveNext
						Loop	
						Rs2.Close
						Set Rs2=nothing
						End If
						%>
                      </select> </td>
                  </tr>
                </table></td>
              <td>&nbsp;</td>
              <td><span class="chamadas">Escolha o T&iacute;tulo<span class="txtmenordourado">*</span><br>
                <img src="../imagens/img_fundo/transp.gif" width="8" height="7"><br>
                </span> <table border="0" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td> <select name="titulo" class="txtmenor" style="width:160px;" >
                        <option selected>Escolha...</option>
                        <%
						Sqltext2 = "SELECT * FROM contato_titulos ORDER BY titulo ASC"
						Set Rs2 = Conn.Execute(Sqltext2)
						If Rs2.EOF AND Rs2.BOF = TRUE Then
						Else
						Do While NOT Rs2.EOF
						%>
                        <option value="<%=Rs2("titulo")%>"><%=Rs2("titulo")%></option>
                        <%
						Rs2.MoveNext
						Loop	
						Rs2.Close
						Set Rs2=nothing
						End If
						%>
                      </select> </td>
                  </tr>
                </table></td>
            </tr>
            <tr valign="bottom"> 
              <td width="158"> <span class="chamadas">Seu Nome<span class="txtmenordourado">*</span><br>
                <img src="../imagens/img_fundo/transp.gif" width="8" height="7"><br>
                </span> <input type="text" name="nome" class="txtmenor" size="30"> 
              </td>
              <td width="3">&nbsp;</td>
              <td width="183" rowspan="3"><span class="chamadas">Sua mensagem<span class="txtmenordourado">*</span></span><br> 
                <img src="../imagens/img_fundo/transp.gif" width="8" height="7"><br> 
                <textarea name="mensagem" class="txtmenor" rows="9" cols="30"></textarea> 
              </td>
            </tr>
            <tr valign="bottom"> 
              <td> <span class="chamadas">Seu E-mail<span class="txtmenordourado">*</span><br>
                <img src="../imagens/img_fundo/transp.gif" width="8" height="7"><br>
                </span> <input type="text" name="email" class="txtmenor" size="30"> 
              </td>
              <td width="3">&nbsp;</td>
            </tr>
            <tr> 
              <td valign="bottom"> <span class="chamadas">Seu Telefone<br>
                <img src="../imagens/img_fundo/transp.gif" width="8" height="7"><br>
                </span> <input type="text" name="ddd" class="txtmenor" size="2" maxlength="2"> 
                <input type="text" name="telefone" class="txtmenor" size="24"> 
              </td>
              <td width="3">&nbsp;</td>
            </tr>
            <tr> 
              <td> <div align="left"> 
                  <input type="image" border="0" name="enviar" src="../imagens/icones/bt_envia.gif" width="40" value="submit" height="16"
                                  onclick="JavaScript:
                                    if (this.form.nome.value != ''){
                                    if (this.form.mensagem.value != ''){
                                    if (this.form.email.value != ''){
                                    if (this.form.singularesrelac.value != ''){
                                      return true;
                                    }else{alert('Campo \'Singular\' não poderá estar em branco.');}
                                    }else{alert('Campo \'Email\' não poderá estar em branco.');}
                                    }else{alert('Campo \'Mensagem\' não poderá estar em branco.');}
                                    }else{alert('Campo \'Nome\' não poderá estar em branco.');form.nome.focus();}
                                    return false;
                                  ;">
                  <img src="../imagens/img_fundo/transp.gif" width="8" height="7"> 
                  <input type="image" border="0" name="limpar" src="../imagens/icones/bt_limpa.gif" width="40" value="reset" height="16"
                                  onclick="JavaScript: this.form.nome.value = this.form.mensagem.value = this.form.email.value = this.form.ddd.value = this.form.telefone.value = ''; return false;" >
                </div></td>
              <td width="3">&nbsp;</td>
              <td class="txtmenordourado" valign="top"> <div align="left">*preenchimento 
                  obrigat&oacute;rio</div></td>
            </tr>
          </table>
        </div>
      </form>
      <!-- -->
    </td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td colspan="2" valign="top"></td>
  </tr>
</table>

<%

If Request.QueryString("mail") = "enviar" Then

	

	If Request.Form("noticias_mail") = "simerro" Then
		noticias_mail = "TRUE"
		'-----------------------------------------------------------------------------------
		' CADASTRO DO NEWSLETTER
		' ----------------------------------------------------------------------------------
		
		'verifica se o e-mail já existe no banco de dados
 		Sql = "SELECT * FROM mail_list WHERE email='" & Request("email") & "'"
	   	Set news = Conn.Execute(Sql)
		'caso o e-mail exista
	   	IF Not news.EOF THEN
		'se o e-mail existe e esta online
		if news("online") = True Then
		' já está cadsatrado não precisa fazer nada
		Else
		' como o e-mail já existe só atualiza o banco de dados
		Sql="Update mail_list SET online=true WHERE ID=" & news("id")
		Conn.Execute(Sql)
		' envia o e-mail de boas vinda para o cidadão
		' necessita o include dos configs de e-mails
		welcome
		' fecha a pergunta de caso o e-mail exista
		end if
		' caso o e-mail não exista no banco de dados
		Else
	   	'cria um tique para o e-mail
	   	randomize timer
		tique = Left(Request("email"),3) & (9876989856 * CInt((RND * 32000) + 100 )) & right(Request("email"),2)
	
		'cadastra o novo usuário
	   Sql = "INSERT INTO mail_list (email,tique,online)"
		  sql = sql + " VALUES('" & Request.Form("email") & "','" & tique & "',true)"
		  Conn.Execute(Sql)
		  ' envia um e-mail de boas vindas para o cidadão
		  welcome
		  ' fechando tudo, só deixando o conn aberto para o registro do outro chamando
		  news.close
		  set news = nothing
		  End If
		'-----------------------------------------------------------------------------------
		' FIM DO CADASTRO DO NEWSLETTER
		' ----------------------------------------------------------------------------------
	Else
		noticias_mail = "FALSE"
	End If
	
	
	If Request.Form("politica_privacidade") = "sim" Then
		politica_privacidade = "TRUE"
	Else
		politica_privacidade = "False"
	End If
	
	

		'SQL = "INSERT INTO mails_site (nome,mail,cidade,estado,titulo,comentario,noticias_mail,politica_privacidade,data)"
		'SQL = SQL + "VALUES ('" & Request("nome") & "','" & Request.form("email") & "','" & Request.form("cidade") & "','" & Request.form("estado") & "',"
		'SQL = SQL + "'" & Request.form("titulo") & "','" & Request.form("comentario") & "'," & noticias_mail & "," & politica_privacidade &","
		'SQL = SQL + "'" & Day(Now) & ("/") & month(Now) & ("/") & Year(Now) & "')"
		
		'Conn.Execute(Sql) 

		mail_site
	
	' Fechando a conexao de dados
	'Conn.Close
	'set Conn = nothing
		
%>
<script>
		msg = "   SEU E-MAIL FOI ENVIADO COM SUCESSO     \n";
		msg += "_____________________________________\n\n";
		msg += "Muito Obrigado <%=Request.form("nome")%> por enviar seu E-mail.\n\n\n";
		msg += "Você acaba de receber um e-mail para de resposta automatica confirmando o nosso recebimento de sua mensagem.\n";
		msg += "Nossa equipe estará respondendo seu e-mail o mais breve possível.\n\n";
		msg += "Abaixo segue os dados enviados ao nosso sistema de mensagens:.\n";
		msg += "=================================================================\n";
		msg += "Seu nome: <%=Request.Form("nome")%>\n";
		msg += "Seu e-mail: <%=Request.Form("mail")%>\n";
		msg += "Cidade: <%=Request.Form("cidade")%> - Uf: <%=Request.Form("estado")%>\n";
		msg += "Título: <%=Request.Form("titulo")%>\n";
		msg += "Comentário: <%=Request.Form("comentario")%>\n";
		msg += "Deseja receber notícias do <%=site_titulo%> por e-mail? - R: <%=Request.Form("noticias_mail")%>\n";
		msg += "Está de acordo com nossa política de privacidade? - R: <%=Request.Form("politica_privacidade")%>\n";
		msg += "=================================================================\n\n\n";
		
		msg += "A equipe do <%=site_titulo%> agradece!!!! Você está sendo redirecionado para nossa página inicial. \n";
		
		alert(msg + "\n\n");
		location.href='default.asp';
		</script>
<%


End If
%>

<!--#include file="rodape.asp"-->