<script language="JavaScript" src="../css/scripts_admin.js"></script>
<!-- #include file="login_nivel2.asp" -->
<!--#include file=".\config_site.inc"-->
<%
if request("tabela") = "Dicas" Then
tabela = "Dicas"
else
tabela = "Noticias"
End If
%>
<div align="center"> <font color="<%=cor_links%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">[<a href="news_default_admin.asp?tabela=<%=tabela%>">voltar 
  para p�gina de not�cias</a>]</font><br>
  <%
If request("editar") = "ok" Then

	Set Rs = Server.CreateObject("ADODB.Recordset")
	Sqltext = "SELECT * FROM "&tabela&" WHERE id = " & request("id") & ""
	Rs.Open SqlText, base_dados
	Session("texto") = Rs("texto")
	titulo = RS("titulo")
	cabeca = Rs("cabeca")
	autor = Rs("autor")
	mailautor = RS("mailautor")
	contagem = RS("contagem")
	recomendado = Rs("recomendado")
	data = Rs("data")
	Rs.Close
	Set Rs = Nothing
	End If
	%>
  <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> </strong></font> </div>
<FORM METHOD="post" ACTION="news_cadastrar_admin.asp?tabela=<%=tabela%>" onSubmit="return CheckForm();" name="frmSendmail">
<input name="contagem" type="hidden" value="<%=contagem%>">
<input name="recomendado" type="hidden" value="<%=recomendado%>">
<input name="data" type="hidden" value="<%=data%>">
<input name="id" type="hidden" value="<%=Request("id")%>">
  <!-- Fim da Tabela de Fundo -->
  <table width="790"   border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center" >
    <tr> 
      <td> 
        <!-- Borda da Tabela -->
        <table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#FFFFFF"  >
          <tr> 
            <td background="../imagens/forum_images/table_bg2.gif"><div align="center"><font color="#336699" size="2" face="Arial, Helvetica, sans-serif"><strong>Editor Interativo 
  para cadastro de <%=tabela%></strong></font></div></td>
          </tr>
          <tr> 
            <td  bgcolor="#999999"></td>
          </tr>
          <tr> 
            <td  bgcolor="#999999"></td>
          </tr>
          <tr> 
            <td bgcolor="#F4F4F4">
<p align="center"> 
                <!-- Inicio da Tabela do Meio -->
                
                <font color="<%=cor_links%>">Todos os campos s&atilde;o obrigat&oacute;rios.</font></strong></font> 
              </div>
              
              <table border="0" align="center" cellpadding="2" cellspacing="1">
                <tr> 
                  <td colspan="2" align="right"><hr align="center" width="80%" noshade color="<%=cor_links%>"></td>
                </tr>
                <tr> 
                  <td width="16%" align="right"><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><span class="titulo" id="var1">Titulo:&nbsp;</span></font></strong></font></td>
                  <td width="84%" nowrap> <font color="<%=cor_links%>" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <input name="titulo" type="text" 
			  style="background-color: #F3F3F3; font-family: Arial; font-size: 8 pt; color: <%=cor_links%>" value="<%=Titulo%>" size="50">
                    </font> </td>
                </tr>
                <tr> 
                  <td align="right"><font color="<%=cor_links%>" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cabe&ccedil;a 
                    da Mat&eacute;ria:<br>
                    </strong>(breve resumo)</font></td>
                  <td valign="top" nowrap><font color="<%=cor_links%>" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <textarea name="cabeca" cols="50" rows="4" id="cabeca" style="background-color: #F3F3F3; font-family: Arial; font-size: 8 pt; color: <%=cor_links%>"><%=cabeca%></textarea>
                    </font></td>
                </tr>
                <tr> 
                  <td align="right"><font color="<%=cor_links%>" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Autor:</strong></font><font color="<%=cor_links%>">&nbsp;</font></td>
                  <td valign="top" nowrap><font color="<%=cor_links%>" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <input name="autor" type="text" id="autor" 
			  style="background-color: #F3F3F3; font-family: Arial; font-size: 8 pt; color: <%=cor_links%>" value="<%=autor%>" size="30">
                    </font> </td>
                </tr>
                <tr> 
                  <td align="right"><font color="<%=cor_links%>" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Email:</strong></font><font color="<%=cor_links%>">&nbsp;</font></td>
                  <td valign="top" nowrap><font color="<%=cor_links%>" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <input name="mailautor" type="text" id="mailautor" 
			 style="background-color: #F3F3F3; font-family: Arial; font-size: 8 pt; color: <%=cor_links%>" value="<%=mailautor%>"size="30">
                    </font></td>
                </tr>
                <tr> 
                  <td colspan="2" align="right"><hr align="center" width="80%" noshade color="<%=cor_links%>"></td>
                </tr>
                <tr> 
                  <td rowspan="2" align="right"><font color="<%=cor_links%>" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Barra 
                    de Ferramentas</strong></font><font color="<%=cor_links%>"><strong>:</strong> 
                    </font></td>
                  <td valign="top" nowrap><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <select name="selectText" 
						 style="background-color: #F3F3F3; font-family: Arial; font-size: 8 pt; color: <%=cor_links%>"
						  onChange="FormatText('FontName', selectText.options[selectText.selectedIndex].value);document.frmSendmail.selectText.options[0].selected = true;" >
                      <option value="0">Fonte</option>
                      <option value="Arial, Helvetica, sans-serif">Arial</option>
                      <option value="Times New Roman, Times, serif">Times</option>
                      <option value="Courier New, Courier, mono">Courier New</option>
                      <option value="Verdana, Arial, Helvetica, sans-serif">Verdana</option>
                    </select>
                    <select name="selectFontSize" 
						  style="background-color: #F3F3F3; font-family: Arial; font-size: 8 pt; color: <%=cor_links%>"
						  onChange="FormatText('FontSize', selectFontSize.options[selectFontSize.selectedIndex].value);document.frmSendmail.selectFontSize.options[0].selected = true;" >
                      <option value="0">Tamanho</option>
                      <option value="1">1</option>
                      <option value="2">2</option>
                      <option value="3">3</option>
                      <option value="4">4</option>
                      <option value="5">5</option>
                      <option value="6">6</option>
                      <option value="7">7</option>
                    </select>
                    <img src="../imagens/botoes_editor/b_visu.gif" width="25" height="24" align="absmiddle" border="0" alt="Ver como est� ficando seu an�ncio" onClick="doPreview('insertorderedlist', '')" style="cursor: hand;"> 
                    <a href="javascript:sobre();"><img src="../imagens/botoes_editor/b_info.gif" width="25" height="24" align="absmiddle" border="0" alt="Sobre o Editor Interativo"  style="cursor: hand;"></a> 
                    <a href="javascript:ajuda();"><img src="../imagens/botoes_editor/b_ajuda.gif" width="25" height="24" align="absmiddle" border="0" alt="Ajuda sobre o que faz cada �cone do Editor"  style="cursor: hand;"></a><br>
                    </font><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../imagens/botoes_editor/b_negrito.gif" width="25" height="24" align="absmiddle" alt="Negrito" onClick="FormatText('bold', '')" style="cursor: hand;"> 
                    <img src="../imagens/botoes_editor/b_italico.gif" width="25" height="24"  align="absmiddle" alt="Italico" onClick="FormatText('italic', '')" style="cursor: hand;"> 
                    <img src="../imagens/botoes_editor/b_sublinhado.gif" width="25" height="24" align="absmiddle" alt="Sublinhado" onClick="FormatText('underline', '')" style="cursor: hand;"> 
                    <img src="../imagens/botoes_editor/b_cor_fonte.gif" width="25" height="24" align="absmiddle" border="0" alt="Setar cor de texto" onClick="foreColor();" style="cursor: hand;"></font></strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    </font><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;<img src="../imagens/botoes_editor/b_alinha_esquerda.gif" width="25" height="24" align="absmiddle" onClick="FormatText('JustifyLeft', '')" style="cursor: hand;" alt="Justificar &agrave; esquerda"> 
                    <img src="../imagens/botoes_editor/b_centralizar.gif" width="25" height="24" align="absmiddle" border="0" alt="Centralizar" onClick="FormatText('JustifyCenter', '')" style="cursor: hand;"> 
                    <img src="../imagens/botoes_editor/b_alinhar_direita.gif" width="25" height="24" align="absmiddle" onClick="FormatText('JustifyRight', '')" style="cursor: hand;" alt="Justificar &agrave; direita"> 
                    <img src="../imagens/botoes_editor/b_alinhar_justificar.gif" width="25" height="24" align="absmiddle" onClick="FormatText('JustifyFull', '')" style="cursor: hand;" alt="Justificar "></font></strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    </font><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;<img src="../imagens/botoes_editor/b_lista_numerada.gif" width="25" height="24" align="absmiddle" border="0" alt="Lista" onClick="FormatText('insertorderedlist', '')" style="cursor: hand;"> 
                    <img src="../imagens/botoes_editor/b_lista.gif" width="25" height="24" align="absmiddle" border="0" alt="Lista" onClick="FormatText('InsertUnorderedList', '')" style="cursor: hand;"></font></strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    </font><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;<img src="../imagens/botoes_editor/b_volta_texto.gif" width="25" height="24" align="absmiddle" onClick="FormatText('Outdent', '')" style="cursor: hand;" alt="Voltar Texto"> 
                    <img src="../imagens/botoes_editor/b_avanca_texto.gif" width="25" height="24" align="absmiddle" border="0" alt="Avan&ccedil;ar texto" onClick="FormatText('indent', '')" style="cursor: hand;"></font></strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    </font></strong></font></td>
                </tr>
                <tr> 
                  <td valign="top" nowrap><p><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../imagens/botoes_editor/b_recortar.gif" width="25" height="24" align="absmiddle" onClick="FormatText('cut')" style="cursor: hand;" alt="Recortar"> 
                      <img src="../imagens/botoes_editor/b_copiar.gif" width="25" height="24" align="absmiddle" onClick="FormatText('copy')" style="cursor: hand;" alt="Copiar"> 
                      <img src="../imagens/botoes_editor/b_cola.gif" width="25" height="24" align="absmiddle" onClick="FormatText('paste')" style="cursor: hand;" alt="Colar">&nbsp; 
                      <img src="../imagens/botoes_editor/b_voltar.gif" width="25" height="24" align="absmiddle" onClick="FormatText('undo', '')" style="cursor: hand;" alt="Justificar &agrave; direita"> 
                      <img src="../imagens/botoes_editor/b_v-refazer.gif" width="25" height="24" align="absmiddle" onClick="FormatText('redo', '')" style="cursor: hand;" alt="Justificar &agrave; direita"> 
                      &nbsp;&nbsp;<img src="../imagens/botoes_editor/b_hiperlink.gif" width="25" height="24" align="absmiddle" border="0" alt="Adicionar Hyperlink" 
								onClick="FormatText('createLink')" style="cursor: hand;"> 
                      <img src="../imagens/botoes_editor/b_mail.gif" width="25" height="24" align="absmiddle" border="0" alt="Adiconar Imagem" onClick="Addmail()" style="cursor: hand;"> 
                      &nbsp;&nbsp;<a href="javascript:upload();"><img src="../imagens/botoes_editor/b_upload_img.gif" width="25" height="24" align="absmiddle" border="0" alt="Envia imagens para o an�ncio"  style="cursor: hand;"></a> 
                      <img src="../imagens/botoes_editor/b_imagem.gif" width="25" height="24" align="absmiddle" border="0" alt="Adiconar Imagem" onClick="AddImage()" style="cursor: hand;"> 
                      &nbsp;&nbsp;<img src="../imagens/botoes_editor/b_tabela.gif" width="25" height="24" align="absmiddle" border="0" alt="Adiconar Tabela" onClick="tableDialog();" style="cursor: hand;"> 
                      <img src="../imagens/botoes_editor/b_linha.gif" width="25" height="24" align="absmiddle" border="0" alt="Adiconar linha" onClick="FormatText('InsertHorizontalRule', '')" style="cursor: hand;"> 
                      &nbsp;&nbsp; 
                  </font></strong></font></p></td>
                </tr>
                <tr> 
                  <td align="right"><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Escreva 
                    aqui o seu texto da not&iacute;cia, use os bot&otilde;es acima 
                    para formatar e tornar seu texto mais agrad&aacute;vel.:</font></strong></font></td>
                  <td valign="top" nowrap> <font color="<%=cor_links%>" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%
'This bit creates a random number to add to the end of the iframe link as IE will cashe the page
'Randomise the system timer
Randomize Timer
%>
                    <script language="javascript">
		    
		    	//Create an iframe and turn on the design mode for it
		    	document.write ('<iframe src="news_textobox.asp?ID=<% = CInt(RND * 2000) %><% If Request("editar") = "ok" Then Response.Write("&mode=show") %>" id="message" width="100%" height="300"></iframe>')
                    	frames.message.document.designMode = "On";
                  
                    </script>
                    <!-- Display a message for IE users with JavaScript turned off -->
                          </font></strong></font><font face="Verdana, Arial, Helvetica, sans-serif">
                          <noscript>
                          <font color="<%=cor_links%>"><strong><font size="1"><br>
                          <br>
                          JavaScript nao esta abilitado no seu navegado, voce 
                          nao pode usar o sistema de editor interativo! </font> 
                          </strong> </font>
                          </noscript>
                          </font> <font color="<%=cor_links%>"><br>
                          <input type="checkbox" onClick="setMode(this.checked)" class='checkbox'>
                          <strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Ver 
                          HTML</font></strong></font></td>
                </tr>
                <tr> 
                  <td align="right">&nbsp;</td>
                  <td valign="top"><font color="<%=cor_links%>"> 
                    <select name="teste" 
						  
						 onChange="Add(teste.options[teste.selectedIndex].value);document.frmSendmail.teste.options[0].selected = true;"
						 style="background-color: #F3F3F3; font-family: Arial; font-size: 8 pt; color: <%=cor_links%>">
                      <%
						  '   onChange="Add('CreateLink', teste.options[teste.selectedIndex].value)" onChange="Add(teste.options[teste.selectedIndex].value)" 'onChange="opa(teste.options[teste.selectedIndex].value)"						   %>> 
                      <option value="0" selected>-- Inserir link de Not�cias relacionadas, 
                      basta selecionar a Not�cia --</option>
                      <%
							 Set Rs = Server.CreateObject("ADODB.Recordset")
			  					Sqltext = "SELECT * FROM noticias WHERE online = TRUE order by id DESC"
								Rs.Open SqlText, base_dados
								If Rs.EOF AND Rs.BOF = TRUE Then
							'nada				
									Else
								Do Until Rs.EOF
								%>
                      <option value="<a href='<%=site_end%>/default.asp?pag=1&noticias=ler&id=<%=Rs("id")%>'><font size=1 face='Verdana, Arial, Helvetica, sans-serif'><br><%=Rs("titulo")%></a>"> 
                      <%=Rs("titulo")%></option>
                      <%
							Rs.MoveNext
							Loop
							End If
							Rs.Close
							set Rs = nothing
							%>
                    </select>
                    <br>
                    </font></td>
                </tr>
                <tr> 
                  <td width="16%" align="right"><font color="<%=cor_links%>"> 
                    <input type="hidden" name="format" value="advHTML">
                    <input type="hidden" name="body" value="">
                    </font></td>
                  <td valign="top"><font color="<%=cor_links%>" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%
			 If Request("editar") = "ok" Then
			 enviar = "Atualizar"
			 Else
			 enviar = "Cadastrar"
			 End If
			 %>
                    <input type="submit" name="Submit" value="<%=enviar%>" 
			  OnClick="document.frmSendmail.body.value = frames.message.document.body.innerHTML;" style="font-family: Arial; font-size: 10 pt; color: <%=cor_links%>">
                    </font></td>
                </tr>
              </table>
              <!-- Fim da Tabela de Fundo -->
              <p align="center"><font color="<%=cor_links%>" size="2" face="Arial, Helvetica, sans-serif"><b> 
                </b></font></p>
              </td>
          </tr>
          <tr> 
            <td  bgcolor="#999999"></td>
          </tr>
          <td> </td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>
<p align="center"><!--#include file="rodape.asp"-->