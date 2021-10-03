<%
Response.Buffer = True
%>
<!-- #include file="login_nivel3.asp" -->
<div align="center">
  <!--#include file=".\config_site.inc"-->
  <script  language="JavaScript">
<!-- Hide from older browsers...

//Function to format text in the text box
function FormatText(command, option){
	
  	frames.message.document.execCommand(command, true, option);
  	frames.message.focus();
}
//Function to add image
function AddImage(){	
	imagePath = prompt('Digite o Endereço da Imagem', 'http://');				
	
	if ((imagePath != null) && (imagePath != "")){					
		frames.message.document.execCommand('InsertImage', false, imagePath);
  		frames.message.focus();
	}
	frames.message.focus();			
}
function cards2(option){	
	imagePath = prompt('Digite o Endereço da Imagem', option);				
	
	if ((imagePath != null) && (imagePath != "")){					
		frames.message.document.execCommand('InsertImage', false, imagePath);
  		frames.message.focus();
	}
	frames.message.focus();			
}
function Add(option){	
	
			frames.message.document.body.innerHTML =  frames.message.document.body.innerHTML + option; 
	
  frames.message.focus();
	}
//Function to add smiley
function cards(option){	
									
	frames.message.document.execCommand('InsertImage', true, option);
  	frames.message.focus();			
}
function AddSmileyIcon(imagePath){	
									
	frames.message.document.execCommand('InsertImage', false, imagePath);
  	frames.message.focus();			
}

//Function to clear form
function ResetForm(){

	if (window.confirm('Tem certeza que deseja limpar todos os campos do e-mail a ser enviado?')){
	 	frames.message.document.body.innerHTML = ''; 
	 	return true;
	 } 
	 return false;		
}

// -->
</script>
  <script  language="JavaScript">
<!-- Hide from older browsers...

//Function to check form is filled in correctly before submitting
function CheckForm() {

	var errorMsg = "";

			
	//Check for the e-mail body
	if (document.frmSendmail.body.value==""){
		errorMsg += "\n\tMensagem\t- Digite a mensagem do E-mail";		
	}	
		
	//If there is aproblem with the form then display an error
	if (errorMsg != ""){
		msg = "_____________________________________________________________\n\n";
		msg += "O seu e-mail não foi enviado por problema(s) em campos em branco.\n";
		msg += "Por Favor, corrija o problema(s) e reenvie o formulário.\n";
		msg += "_____________________________________________________________\n\n";
		msg += "Os seguintes campo(s) não estão corretos: -\n";
		
		errorMsg += alert(msg + errorMsg + "\n\n");
		return false;
	}
	
	return true;
}
function tableDialog()

{
   //----- Creates A Table Dialog And Passes Values To createTable() -----
   var rtNumRows = null;
   var rtNumCols = null;
   var rtTblAlign = null;
   var rtTblWidth = null;
   var rtspacing = null;
   var rtpadding = null;
   var rtbrdsize = null;
   showModalDialog("table_news.htm",window,"status:false;dialogWidth:320px;dialogHeight:310px");
}

function createTable()
{

   //----- Criando uma tabela definida pelo usuário -----
 
   var cursor = frames.message.document.selection.createRange();
   if (rtNumRows == "" || rtNumRows == "0")
   {
      rtNumRows = "1";
   }
   if (rtNumCols == "" || rtNumCols == "0")
   {
      rtNumCols = "1";
   }
   var rttrnum=1
   var rttdnum=1
   var rtNewTable = "<table border='" + rtbrdsize + "' align='" + rtTblAlign + "' cellpadding='" + rtpadding + "' cellspacing='" + rtspacing + "' width='" + rtTblWidth + "'>"
   while (rttrnum <= rtNumRows)
   {
      rttrnum=rttrnum+1
      rtNewTable = rtNewTable + "<tr>"
      while (rttdnum <= rtNumCols)
      {
         rtNewTable = rtNewTable + "<td>&nbsp;</td>"
         rttdnum=rttdnum+1
      }
      rttdnum=1
      rtNewTable = rtNewTable + "</tr>"
   }
   rtNewTable = rtNewTable + "</table>"
   frames.message.focus();
   cursor.pasteHTML(rtNewTable);
}
function setMode(bMode) {

	var sTmp;

  	isHTMLMode = bMode;

  	if (isHTMLMode) {

		sTmp=frames.message.document.body.innerHTML;

		frames.message.document.body.innerText=sTmp;

		

	} else {

		sTmp=frames.message.document.body.innerText;

		frames.message.document.body.innerHTML=sTmp;

		

	}

  frames.message.focus();

}
function foreColor()	{

	var arr = showModalDialog("select_color.html",window,"font-family:Verdana; font-size:12; status:false; dialogWidth:22; dialogHeight:21" );

	if (arr != null) cmdExec("ForeColor",arr);	

}

function cmdExec(cmd,opt) {

  	frames.message.document.execCommand(cmd,"",opt);

	frames.message.focus();

}

function doPreview(){

     temp = frames.message.document.body.innerHTML;

     preWindow= open('', 'previewWindow', 'width=500,height=440,status=yes,scrollbars=yes,resizable=yes,toolbar=no,menubar=yes');

     preWindow.document.open();

     preWindow.document.write(temp);

     preWindow.document.close();

}

// -->
</script>
  <br>
  <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="newsletter_admin_menu.asp">Voltar 
  para página administradora de Newsletter</a></font> </div>
<p></p>
<% 	If Request("send_voce") = "ok" Then	%>
<div align="center"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Um 
  E-mail com um Preview para o e-mail <%=admin_mail%> acaba de ser enviado. 
  <%	 End If 	%>
</font></strong> </div>
<% If request("form") = "ver" Then	%>
<form method=post name="frmSendmail" action="?form=envia" onSubmit="return CheckForm();" onReset="return ResetForm();">
<table width="190" border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center">
            <tr> 
              <td> <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
                  <tr > 
                    <td height="23" background="../imagens/icos_e_fundos/table_bg.gif" class="f_vermelho_3_b"> 
                      <div align="center">Sistema de Newsletter (Envio de Notícias por E-mail)</div></td>
                  </tr>
                  <tr bgcolor="#EEEEEE"> 
                    <td bgcolor="#EEEEEE">
                    
                   
 <table width="100%" height="139" border="0" align="center" background="../imagens/fundo_mail/fundo_mail_r2_c2.gif" >
                      <tr> 
                        <td colspan="2" height="30" class="text" align="left"><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">*Indica 
                          campos obrigat&oacute;rios</font></strong></font></td>
                      </tr>
                      <tr > 
                        <td align="right" height="12"><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Not&iacute;cias:</font></strong></font></td>
                        <td height="12"><font color="<%=cor_links%>">
                          <select name="teste2" 
						  
						 onChange="Add(teste2.options[teste2.selectedIndex].value);document.frmSendmail.teste2.options[0].selected = true;"
						 >
                            <%
						  '   onChange="Add('CreateLink', teste.options[teste.selectedIndex].value)" onChange="Add(teste.options[teste.selectedIndex].value)" 'onChange="opa(teste.options[teste.selectedIndex].value)"						   %>> 
                            <option value="0" selected>-- Escolha um Link de Notícias 
                            --</option>
                            <%
							 Set Rs = Server.CreateObject("ADODB.Recordset")
			  					Sqltext = "SELECT * FROM noticias WHERE online = TRUE order by id DESC"
								Rs.Open SqlText, base_dados
								If Rs.EOF AND Rs.BOF = TRUE Then
							'nada				
									Else
								Do Until Rs.EOF
								%>
                            <option value="<li><a href='<%=site_end%>/noticias_ler.asp?id=<%=Rs("id")%>&noticias=ler&tabela=noticias'><font size=1 face='Verdana, Arial, Helvetica, sans-serif'><%=Rs("titulo")%></a> - <%=Rs("cabeca")%> ...</font> </li><br>"> 
                            <%=Rs("titulo")%></option>
                            <%
							Rs.MoveNext
							Loop
							End If
							Rs.Close
							set Rs = nothing
							%>
                          </select>
                          </font></td>
                      </tr>
                      <tr> 
                      
					  
					  
					    <td valign="middle" align="right" height="22" width="15%"><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><span class="text">Formato 
                          de Textot:</span></font></strong></font></td>
                        <td height="22" width="85%" valign="bottom"> <table width="100%" border="0" cellspacing="0" cellpadding="1">
                            <tr> 
                              <td> <font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                                <select name="selectText" 
					
						  onChange="FormatText('FontName', selectText.options[selectText.selectedIndex].value);document.frmSendmail.selectText.options[0].selected = true;" >
                                  <option value="0">-- Tipo de Fonte --</option>
                                  <option value="Arial, Helvetica, sans-serif">Arial</option>
                                  <option value="Times New Roman, Times, serif">Times</option>
                                  <option value="Courier New, Courier, mono">Courier 
                                  New</option>
                                  <option value="Verdana, Arial, Helvetica, sans-serif">Verdana</option>
                                </select>
                                <select name="selectFontSize" 
						  
						  onChange="FormatText('FontSize', selectFontSize.options[selectFontSize.selectedIndex].value);document.frmSendmail.selectFontSize.options[0].selected = true;" >
                                  <option value="0">-- Tamanhos de Fonte--</option>
                                  <option value="1">1</option>
                                  <option value="2">2</option>
                                  <option value="3">3</option>
                                  <option value="4">4</option>
                                  <option value="5">5</option>
                                  <option value="6">6</option>
                                  <option value="7">7</option>
                                </select>
                                <img src="../imagens/botoes_editor/b_negrito.gif" width="25" height="24" align="absmiddle" alt="Negrito" onClick="FormatText('bold', '')" style="cursor: hand;"> 
                                <img src="../imagens/botoes_editor/b_italico.gif" width="25" height="24"  align="absmiddle" alt="Italico" onClick="FormatText('italic', '')" style="cursor: hand;"> 
                                <img src="../imagens/botoes_editor/b_sublinhado.gif" width="25" height="24" align="absmiddle" alt="Sublinhado" onClick="FormatText('underline', '')" style="cursor: hand;"> 
                                <img src="../imagens/botoes_editor/b_cor_fonte.gif" width="25" height="24" align="absmiddle" border="0" alt="Adiconar Tabela" onClick="foreColor();" style="cursor: hand;"></font></strong></font></td>
                            </tr>
                            <tr> 
                              <td><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../imagens/botoes_editor/b_recortar.gif" width="25" height="24" align="absmiddle" onClick="FormatText('cut')" style="cursor: hand;" alt="Recortar"> 
                                <img src="../imagens/botoes_editor/b_copiar.gif" width="25" height="24" align="absmiddle" onClick="FormatText('copy')" style="cursor: hand;" alt="Copiar"> 
                                <img src="../imagens/botoes_editor/b_cola.gif" width="25" height="24" align="absmiddle" onClick="FormatText('paste')" style="cursor: hand;" alt="Colar"> 
                                <img src="../imagens/botoes_editor/b_voltar.gif" width="25" height="24" align="absmiddle" onClick="FormatText('undo', '')" style="cursor: hand;" alt="Justificar &agrave; direita"> 
                                <img src="../imagens/botoes_editor/b_v-refazer.gif" width="25" height="24" align="absmiddle" onClick="FormatText('redo', '')" style="cursor: hand;" alt="Justificar &agrave; direita"> 
                                <img src="../imagens/botoes_editor/b_alinha_esquerda.gif" width="25" height="24" align="absmiddle" onClick="FormatText('JustifyLeft', '')" style="cursor: hand;" alt="Justificar &agrave; esquerda"> 
                                <img src="../imagens/botoes_editor/b_centralizar.gif" width="25" height="24" align="absmiddle" border="0" alt="Centralizar" onClick="FormatText('JustifyCenter', '')" style="cursor: hand;"> 
                                <img src="../imagens/botoes_editor/b_alinhar_direita.gif" width="25" height="24" align="absmiddle" onClick="FormatText('JustifyRight', '')" style="cursor: hand;" alt="Justificar &agrave; direita"> 
                                <img src="../imagens/botoes_editor/b_alinhar_justificar.gif" width="25" height="24" align="absmiddle" onClick="FormatText('JustifyFull', '')" style="cursor: hand;" alt="Justificar "> 
                                <img src="../imagens/botoes_editor/b_lista_numerada.gif" width="25" height="24" align="absmiddle" border="0" alt="Lista" onClick="FormatText('insertorderedlist', '')" style="cursor: hand;"> 
                                <img src="../imagens/botoes_editor/b_lista.gif" width="25" height="24" align="absmiddle" border="0" alt="Lista" onClick="FormatText('InsertUnorderedList', '')" style="cursor: hand;"> 
                                <img src="../imagens/botoes_editor/b_volta_texto.gif" width="25" height="24" align="absmiddle" onClick="FormatText('Outdent', '')" style="cursor: hand;" alt="Voltar Texto"> 
                                <img src="../imagens/botoes_editor/b_avanca_texto.gif" width="25" height="24" align="absmiddle" border="0" alt="Avan&ccedil;ar texto" onClick="FormatText('indent', '')" style="cursor: hand;"> 
                                <img src="../imagens/botoes_editor/b_hiperlink.gif" width="25" height="24" align="absmiddle" border="0" alt="Adicionar Hyperlink" 
								onClick="FormatText('createLink')" style="cursor: hand;"> 
                                <img src="../imagens/botoes_editor/b_imagem.gif" width="25" height="24" align="absmiddle" border="0" alt="Adiconar Imagem" onClick="AddImage()" style="cursor: hand;"> 
                                <img src="../imagens/botoes_editor/b_tabela.gif" width="25" height="24" align="absmiddle" border="0" alt="Adiconar Tabela" onClick="tableDialog();" style="cursor: hand;"> 
                                <img src="../imagens/botoes_editor/b_linha.gif" width="25" height="24" align="absmiddle" border="0" alt="Adiconar linha" onClick="FormatText('InsertHorizontalRule', '')" style="cursor: hand;"> 
                                </font></strong></font></td>
                          </tr>
                          </table></td>
                      </tr>
                      <tr > 
                        <td valign="top" align="right" height="61" width="15%" ><font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><span class="text">Mensagem*:</span></font></strong></font></td>
                        <td height="61" width="85%" valign="top"> <font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%
'This bit creates a random number to add to the end of the iframe link as IE will cashe the page
'Randomise the system timer
Randomize Timer
%>
                          <script language="javascript">
		    
		    	//Create an iframe and turn on the design mode for it
		    	document.write ('<iframe src="newsletter_textobox.asp?ID=<% = CInt(RND * 2000) %><% If Request("send_voce") = "ok" Then Response.Write("&mode=show") %>" id="message" width="510" height="200"></iframe>')
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
                      <td valign="top" align="right" height="2" width="15%" > 
                        <font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                        <input type="hidden" name="format" value="advHTML">
                        <input type="hidden" name="body" value="">
                        </font></strong></font></td>
                      <td height="2" width="85%" align="left"> <font color="<%=cor_links%>"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                        <input name="Submit" type="submit" class="f_vermelho_2" 
				   OnClick="document.frmSendmail.body.value = frames.message.document.body.innerHTML;" value="Enviar um Preview para você">
                      <input name="Submit" type="submit" class="f_vermelho_2" 
				  OnClick="document.frmSendmail.body.value = frames.message.document.body.innerHTML;" value="Enviar para todos os Membros">
                      <input name="Reset" type="reset" class="f_vermelho_2" value="Limpar Formulário" >
                        </font></strong></font></td>
                      </tr>
                  </table>
</form>
 
                    
                    </td>
                  </tr>
                </table></td>
            </tr>
          </table>
<% End If %>
<% if request("form") = "envia" Then 	%>
<!--#include file="newsletter_envia.asp"-->
<% End If %>
<p align="center"><!--#include file="rodape.asp"-->