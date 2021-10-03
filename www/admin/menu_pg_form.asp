<!-- #include file="login_nivel3.asp" -->

<script language="JavaScript" src="../functions/scripts_admin.js"></script>

<!--#include file=".\config_site.inc"-->
<%
if request("ac") = "cadastrar" Then
		pe = "SELECT * FROM menu_pgs WHERE id_menu = "&request("id_menu")
		Set pg = Conn.Execute(pe)
		IF Not pg.EOF THEN
			response.Redirect "?ac=editar&id_menu="&request("id_menu")
		pg.close
		End If
End If
		%>
<div align="center"> <font color="<%=cor_links%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">[<a href="menus_default.asp">voltar 
  para página de Menus</a>]</font><br>
  <%
If request("ac") = "editar" Then

	Set Rs = Server.CreateObject("ADODB.Recordset")
	Sqltext = "SELECT * FROM menu_pgs WHERE id_menu = " & request("id_menu") & ""
	Rs.Open SqlText, base_dados
	Session("texto") = Rs("conteudo")
	titulo = RS("titulo")
	subtitulo= Rs("subtitulo")
	Rs.Close
	Set Rs = Nothing
	End If
	%>
  <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> </strong></font> </div>
<FORM METHOD="post" ACTION="menu_pg_cadastrar.asp?ac=<%=request("ac")%>" onSubmit="return CheckForm();" name="frmSendmail">

<input name="data" type="hidden" value="<%=data%>">
<input name="id_menu" type="hidden" value="<%=Request("id_menu")%>">
  <!-- Fim da Tabela de Fundo -->
  <table width="790"   border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center" >
    <tr> 
      <td> 
        <!-- Borda da Tabela -->
        <table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#FFFFFF"  >
          <tr> 
            <td background="../imagens/forum_images/table_bg2.gif"><div align="center"><font color="#336699" size="2" face="Arial, Helvetica, sans-serif"><strong>Editor 
                Interativo para cadastro de P&aacute;ginas</strong></font></div></td>
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
			  value="<%=Titulo%>" size="50">
                    </font> </td>
                </tr>
                <tr> 
                  <td align="right"><font color="<%=cor_links%>" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Sub-Titulo</strong></font></td>
                  <td valign="top" nowrap><font color="<%=cor_links%>" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <input name="subtitulo" type="text" 
			  value="<%=subtitulo%>" size="50">
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
                    <img src="../imagens/botoes_editor/b_corretor.gif" width="25" height="24" align="absmiddle" border="0" alt="Ver como está ficando seu anúncio" onClick="doPreview('insertorderedlist', '')" style="cursor: hand;"> 
                    <a href="javascript:sobre();"><img src="../imagens/botoes_editor/b_info_editor.gif" width="25" height="24" align="absmiddle" border="0" alt="Sobre o Editor Interativo"  style="cursor: hand;"></a> 
                    <a href="javascript:ajuda();"><img src="../imagens/botoes_editor/b_ajuda.gif" width="25" height="24" align="absmiddle" border="0" alt="Ajuda sobre o que faz cada ícone do Editor"  style="cursor: hand;"></a><br>
                    </font><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../imagens/botoes_editor/b_negrito.gif" width="25" height="24" align="absmiddle" alt="Negrito" onClick="FormatText('bold', '')" style="cursor: hand;"><img src="../imagens/botoes_editor/b_italico.gif" width="25" height="24"  align="absmiddle" alt="Italico" onClick="FormatText('italic', '')" style="cursor: hand;"><img src="../imagens/botoes_editor/b_sublinhado.gif" width="25" height="24" align="absmiddle" alt="Sublinhado" onClick="FormatText('underline', '')" style="cursor: hand;"><img src="../imagens/botoes_editor/b_cor_fonte.gif" width="25" height="24" align="absmiddle" border="0" alt="Setar cor de texto" onClick="foreColor();" style="cursor: hand;"></font></strong></font><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../imagens/botoes_editor/b_alinha_esquerda.gif" width="25" height="24" align="absmiddle" onClick="FormatText('JustifyLeft', '')" style="cursor: hand;" alt="Justificar &agrave; esquerda"><img src="../imagens/botoes_editor/b_centralizar.gif" width="25" height="24" align="absmiddle" border="0" alt="Centralizar" onClick="FormatText('JustifyCenter', '')" style="cursor: hand;"><img src="../imagens/botoes_editor/b_alinhar_direita.gif" width="25" height="24" align="absmiddle" onClick="FormatText('JustifyRight', '')" style="cursor: hand;" alt="Justificar &agrave; direita"><img src="../imagens/botoes_editor/b_alinhar_justificar.gif" width="25" height="24" align="absmiddle" onClick="FormatText('JustifyFull', '')" style="cursor: hand;" alt="Justificar "></font></strong></font><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../imagens/botoes_editor/b_lista_numerada.gif" width="25" height="24" align="absmiddle" border="0" alt="Lista" onClick="FormatText('insertorderedlist', '')" style="cursor: hand;"><img src="../imagens/botoes_editor/b_lista.gif" width="25" height="24" align="absmiddle" border="0" alt="Lista" onClick="FormatText('InsertUnorderedList', '')" style="cursor: hand;"></font></strong></font><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../imagens/botoes_editor/b_volta_texto.gif" width="25" height="24" align="absmiddle" onClick="FormatText('Outdent', '')" style="cursor: hand;" alt="Voltar Texto"><img src="../imagens/botoes_editor/b_avanca_texto.gif" width="25" height="24" align="absmiddle" border="0" alt="Avan&ccedil;ar texto" onClick="FormatText('indent', '')" style="cursor: hand;"></font></strong></font><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../imagens/botoes_editor/b_recortar.gif" width="25" height="24" align="absmiddle" onClick="FormatText('cut')" style="cursor: hand;" alt="Recortar"><img src="../imagens/botoes_editor/b_copiar.gif" width="25" height="24" align="absmiddle" onClick="FormatText('copy')" style="cursor: hand;" alt="Copiar"><img src="../imagens/botoes_editor/b_cola.gif" width="25" height="24" align="absmiddle" onClick="FormatText('paste')" style="cursor: hand;" alt="Colar"><img src="../imagens/botoes_editor/b_voltar.gif" width="25" height="24" align="absmiddle" onClick="FormatText('undo', '')" style="cursor: hand;" alt="Justificar &agrave; direita"><img src="../imagens/botoes_editor/b_v-refazer.gif" width="25" height="24" align="absmiddle" onClick="FormatText('redo', '')" style="cursor: hand;" alt="Justificar &agrave; direita"><img src="../imagens/botoes_editor/b_hiperlink.gif" width="25" height="24" align="absmiddle" border="0" alt="Adicionar Hyperlink" 
								onClick="FormatText('createLink')" style="cursor: hand;"><img src="../imagens/botoes_editor/b_mail.gif" width="25" height="24" align="absmiddle" border="0" alt="Adiconar Imagem" onClick="Addmail()" style="cursor: hand;"><a href="javascript:upload();"><img src="../imagens/botoes_editor/b_imagem_upload2.gif" width="25" height="24" align="absmiddle" border="0" alt="Envia imagens para o anúncio"  style="cursor: hand;"></a><img src="../imagens/botoes_editor/b_imagem.gif" width="25" height="24" align="absmiddle" border="0" alt="Adiconar Imagem" onClick="AddImage()" style="cursor: hand;"><img src="../imagens/botoes_editor/b_tabela.gif" width="25" height="24" align="absmiddle" border="0" alt="Adiconar Tabela" onClick="tableDialog();" style="cursor: hand;"><img src="../imagens/botoes_editor/b_linha.gif" width="25" height="24" align="absmiddle" border="0" alt="Adiconar linha" onClick="FormatText('InsertHorizontalRule', '')" style="cursor: hand;"><img src="../imagens/botoes_editor/b_mostra_html.gif" width="25" height="24" onClick="setMode(this.checked)" style="cursor: hand;"  align="absmiddle"> 
                    </font></strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    </font></strong></font></td>
                </tr>
                <tr> 
                  <td valign="top" nowrap><p><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      </font></strong></font></p></td>
                </tr>
                <tr> 
                  <td align="right"><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Escreva 
                    aqui o conte&uacute;do dessa p&aacute;gina, use os bot&otilde;es 
                    acima para formatar e tornar seu texto mais agrad&aacute;vel.:</font></strong></font></td>
                  <td valign="top" nowrap> 
                    <%
'This bit creates a random number to add to the end of the iframe link as IE will cashe the page
'Randomise the system timer
Randomize Timer
%>
                    <script language="javascript">
		    
		    	//Create an iframe and turn on the design mode for it
		    	document.write ('<iframe src="news_textobox.asp?ID=<% = CInt(RND * 2000) %><% If Request("ac") = "editar" Then Response.Write("&mode=show") %>" id="message" width="100%" height="300"></iframe>')
                    	frames.message.document.designMode = "On";
                  
                    </script> </td>
                </tr>
                <tr> 
                  <td align="right">    	<font color="<%=cor_links%>">
                    	<input type="checkbox" onClick="setMode(this.checked)" class='checkbox'>
                        <strong>
                        	<font size="1" face="Verdana, Arial, Helvetica, sans-serif">
                            	Ver HTML
                            </font>
                        </strong>
                    </font></td>
                  <td valign="top"><font color="<%=cor_links%>">&nbsp; </font></td>
                </tr>
                <tr> 
                  <td width="16%" align="right"><font color="<%=cor_links%>"> 
                    <input type="hidden" name="format" value="advHTML">
                    <input type="hidden" name="body" value="">
                    </font></td>
                  <td valign="top"><font color="<%=cor_links%>" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%
			 If Request("ac") = "editar" Then
			 enviar = "Atualizar"
			 Else
			 enviar = "Cadastrar"
			 End If
			 %>
                    <input type="submit" name="Submit" value="<%=enviar%>" 
			  OnClick="document.frmSendmail.body.value = frames.message.document.body.innerHTML;" style="font-family: Arial; font-size: 8 pt; color: <%=cor_links%>">
                    </font></td>
                </tr>
              </table>
              <!-- Fim da Tabela de Fundo -->
              
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
<!-- #include file="rodape.asp" -->
