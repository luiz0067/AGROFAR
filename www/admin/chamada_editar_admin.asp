<script language="JavaScript" src="../css/scripts_admin.js"></script>
<!-- #include file="login_nivel3.asp" -->
<!--#include file=".\config_site.inc"-->
<%	if request("msg") = "1" Then	%>
<table width="390" border="0" cellspacing="1" cellpadding="1" align="center" class="moldura1">
  <tr>
    <td><div align="center"><img src="../imagens/icos_e_fundos/warning.png" align="absmiddle"> <span style="color: #FF0000; font-weight: bold">
    Sua Pagina foi Salva com Sucesso!</span> </div></td>
  </tr>
</table>
<br />
    <%	End If	%>
  <%
  
  Set Rs = Server.CreateObject("ADODB.Recordset")
	Sqltext = "SELECT * FROM chamada WHERE id = " & request("id") & ""
	Rs.Open SqlText, base_dados
	Session("texto") = Rs("chamada")
	Rs.Close
	Set Rs = Nothing
	
If request("editar") = "ok" Then
msg = Request("Body")
msg = replace(msg,"|_aspas_|","'")
msg = replace(msg,"'","''") 

Sql="Update chamada SET chamada = '" & msg & "' WHERE ID=" & request("id")
		Conn.Execute(Sql)	   
response.Redirect("chamada_default_admin.asp?msg=2")
End If


	%>
  <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> </strong></font> </div>
<FORM METHOD="post" ACTION="chamada_editar_admin.asp?editar=ok&id=<%=request("id")%>" onSubmit="return CheckForm();" name="frmSendmail">
  <!-- Fim da Tabela de Fundo -->
  <table width="790"   border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center" >
    <tr> 
      <td> 
        <!-- Borda da Tabela -->
        <table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#FFFFFF"  >
          <tr> 
            <td background="../imagens/icos_e_fundos/table_bg2.gif"><div align="center"><font color="#336699" size="2" face="Arial, Helvetica, sans-serif"><strong>Editor Interativo 
  de Chamadas na P&aacute;gina Inicial</strong></font></div></td>
          </tr>
          <tr> 
            <td  bgcolor="#999999"></td>
          </tr>
          <tr> 
            <td  bgcolor="#999999"></td>
          </tr>
          <tr> 
            <td bgcolor="#F4F4F4">

              <table border="0" align="center" cellpadding="2" cellspacing="1">
             
                  <td rowspan="2" align="right"><font color="<%=cor_links%>" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Barra 
                    de Ferramentas</strong></font><font color="<%=cor_links%>"><strong>:</strong> 
                    </font></td>
                  <td width="84%" valign="top" nowrap><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
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
                    <img src="../imagens/outros_icos/sm-gear.gif" width="20" height="20" align="absmiddle" border="0" alt="Ver como está ficando seu anúncio" onClick="doPreview('insertorderedlist', '')" style="cursor: hand;"> 
                    <a href="javascript:sobre();"><img src="../imagens/icones/info.png" width="16" height="16" align="absmiddle" border="0" alt="Sobre o Editor Interativo"  style="cursor: hand;"></a> 
                    <a href="javascript:ajuda();"><img src="../imagens/icos_e_fundos/question.png" width="16" height="16" align="absmiddle" border="0" alt="Ajuda sobre o que faz cada ícone do Editor"  style="cursor: hand;"></a><br>
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
                      &nbsp;&nbsp;<a href="javascript:upload();"><img src="../imagens/botoes_editor/b_imagem_upload2.gif" width="25" height="24" align="absmiddle" border="0" alt="Envia imagens para o anúncio"  style="cursor: hand;"></a> 
                      <img src="../imagens/botoes_editor/b_imagem.gif" width="25" height="24" align="absmiddle" border="0" alt="Adiconar Imagem" onClick="AddImage()" style="cursor: hand;"> 
                      &nbsp;&nbsp;<img src="../imagens/botoes_editor/b_tabela.gif" width="25" height="24" align="absmiddle" border="0" alt="Adiconar Tabela" onClick="tableDialog();" style="cursor: hand;"> 
                      <img src="../imagens/botoes_editor/b_linha.gif" width="25" height="24" align="absmiddle" border="0" alt="Adiconar linha" onClick="FormatText('InsertHorizontalRule', '')" style="cursor: hand;"> 
                      &nbsp;&nbsp; 
                      </font></strong></font></p></td>
                </tr>
                <tr> 
                  <td align="right"><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Escreva aqui tudo que desejar e clique em Salvar:</font></strong></font></td>
                  <td valign="top" nowrap> <font color="<%=cor_links%>" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%
'This bit creates a random number to add to the end of the iframe link as IE will cashe the page
'Randomise the system timer
Randomize Timer
%>
                    <script language="javascript">
		    
		    	//Create an iframe and turn on the design mode for it
		    	document.write ('<iframe src="news_textobox.asp?ID=<% = CInt(RND * 2000) %><%	Response.Write("&mode=show")	%>" id="message" width="90%" height="150"></iframe>')
                    	frames.message.document.designMode = "On";
                  
                    </script>
                    <!-- Display a message for IE users with JavaScript turned off -->
                          </font></strong></font><font face="Verdana, Arial, Helvetica, sans-serif">
                          <noscript>
                          <font color="<%=cor_links%>"><strong><font size="1"><br>
                          <br>
                          JavaScript não esta abilitado no seu navegado, você 
                          não pode usar o sistema de editor interativo! </font> 
                          </strong> </font>
                          </noscript>
                          </font> <font color="<%=cor_links%>"><br>
                          <input type="checkbox" onClick="setMode(this.checked)" class='checkbox'>
                          <strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Ver 
                          HTML</font></strong></font></td>
                </tr>
                <tr> 
                  <td width="16%" align="right"><font color="<%=cor_links%>"> 
                    <input type="hidden" name="format" value="advHTML">
                    <input type="hidden" name="body" value="">
                    </font></td>
                  <td valign="top"><font color="<%=cor_links%>" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    
                    <input type="submit" name="Submit" value="Atualizar" 
			  OnClick="document.frmSendmail.body.value = frames.message.document.body.innerHTML;" style="font-family: Arial; font-size: 10 pt; color: <%=cor_links%>">
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
<p align="center"><!--#include file="rodape.asp"-->