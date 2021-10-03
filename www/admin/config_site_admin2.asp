<!-- #include file="login_nivel3.asp" -->
<!--#include file=".\config_site.inc"-->
<script>
function foreColor()	{
	var arr = showModalDialog("select_color.html",window,"font-family:Verdana; font-size:12; status:false; dialogWidth:22; dialogHeight:21" );
	if (arr != null) cmdExec("ForeColor",arr);	
}
function cmdExec(cmd,opt) {
	document.config.cor_links.value=(cmd,"","#" + opt);
	document.all.cor_link.style.color=(cmd,"","#" + opt);
	config.cor_links.focus();
}
function foreColor2()	{
	var arr = showModalDialog("select_color.html",window,"font-family:Verdana; font-size:12; status:false; dialogWidth:22; dialogHeight:21" );
	if (arr != null) cmdExec2("ForeColor",arr);	
}
function cmdExec2(cmd,opt) {
	document.config.cor_fundo_box.value=(cmd,"","#" + opt);
	document.config.ex_cor_fundo_box.style.backgroundColor=(cmd,"","#" + opt);
	
	
	//document.ColorPreview.style.backgroundColor = '#' + opt;
  //document.all.ColorHex.value = '#' + color;
  //if (document.all) {
  document.body.style.scrollbarBaseColor = (cmd,"","#" + opt);
  //}
	//config.cor_fundo_box.focus();

}
function fontesete(opt) {
	document.config.site_fonte.value=(opt);
	
	//document.config.tipo_fonte.style.fontFamily= (opt) ;
	document.all.tipo_fonte2.style.fontFamily= (opt) ;
	
	config.site_fonte.focus();
}
function upload()	{
	window.open('config_site_upload.asp','WinExa','toolbar=0,location=0,status=1,menubar=0,scrollbars=0,resizable=0,width=400,height=150')
}
function foreColor3()	{
	var arr = showModalDialog("select_color.html",window,"font-family:Verdana; font-size:12; status:false; dialogWidth:22; dialogHeight:21" );
	if (arr != null) cmdExec3("ForeColor",arr);	
}
function cmdExec3(cmd,opt) {
	document.config.cor_fundo_site.value=(cmd,"","#" + opt);
	document.all.m_cor_fundo_site.style.backgroundColor=(cmd,"","#" + opt);
	config.cor_fundo_site.focus();
}	
</script>
<%
		' Abre o banco de dados e seleciona o id que quer ser lido
		set Rs = server.createobject("ADODB.recordset")
		Sqltext = "SELECT * FROM configuracoes_site"
		Set Rs = Conn.Execute(Sqltext)
%>

<FORM ACTION="?config_site=atualizar" METHOD="POST" name="config" id="config" >
  
  <table width="650" border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center">
  <tr>
    <td>
	
	
  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
          <tr > 
            <td height="23" colspan="4" background="../imagens/forum_images/table_bg.gif"> 
              <center>
                <strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">:: 
                Configura&ccedil;&otilde;es do Site::</font> </strong> </center></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td colspan="4" bgcolor="#EEEEEE"><div align="center"><font color="#FF0000" size="1" face="Arial, Helvetica, sans-serif">Strings 
                do site, s&oacute; altere esses campos com o consentimento do 
                Webmaster do Site.</font></div></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td width="23%" bgcolor="#EEEEEE"> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;Titulo 
                do Site: </font></div></td>
            <td width="37%"> <font size="1" face="Arial, Helvetica, sans-serif"> 
              <input name="site_titulo" type="text" id="site_titulo"  class="text" style="width:195px;" value="<%= Rs("site_titulo")%>" >
              </font> </td>
            <td width="5%"><div align="right"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1"></font></font></div></td>
            <td width="35%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td width="23%" height="24" bgcolor="#EEEEEE"> <div align="right" background="../imagens/forum_images/table_bg2.gif"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Endere&ccedil;o 
                do Site:</font></div></td>
            <td width="37%"><font size="1" face="Arial, Helvetica, sans-serif"> 
              <input name="site_end" type="text"  class="text" id="site_end" style="width:195px;" value="<%= Rs("site_end")%>">
              </font></td>
            <td colspan="2" rowspan="3"> <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1"></font></font></div>
              <font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font> 
              <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1"></font></font></div>
              <div align="left"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font> 
                <font face="Verdana, Arial, Helvetica, sans-serif"><font size="1"></font></font></div>
              <div align="left"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;
                <input name="logo_model" type="image" src="../<%=logo_empresa%>" border="1" value="">
                <br>
                <font color="#FF0000">Ex: do campo Logo Empresa</font></font></div></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td width="23%" bgcolor="#EEEEEE"> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">E-mail 
                do Site:</font></div></td>
            <td width="37%"><font size="1" face="Arial, Helvetica, sans-serif"> 
              <input name="admin_mail" type="text"  class="text" id="admin_mail" style="width:195px;" value="<%= Rs("admin_mail")%>">
              </font></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td width="23%" bgcolor="#EEEEEE"> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Host 
                Mail :</font></div></td>
            <td width="37%"><font size="1" face="Arial, Helvetica, sans-serif"> 
              <input name="host_mail" type="text"  class="text" id="host_mail" style="width:195px;" value="<%= Rs("host_mail")%>">
              </font></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td width="23%" bgcolor="#EEEEEE"> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Logo 
                da Empresa:</font></div></td>
            <td width="37%"> <input name="logo_empresa" type="text"  class="text" id="logo_empresa" style="width:195px;" value="<%= Rs("logo_empresa")%>" > 
              <img src="../imagens/forum_images/post_button_image_upload_nova.gif" width="25" height="24" onClick="upload();" style="cursor: hand;"> 
            </td>
            <td colspan="2"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif">Clique 
              para enviar o seu logo</font></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td width="23%" bgcolor="#EEEEEE"> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Nome 
                Dados: </font></div></td>
            <td> <input name="dados" type="text"  class="text" id="dados" style="width:195px;" value="<%= Rs("dados")%>" disabled > 
              <input name="dados" type="hidden" value="<%= Rs("dados")%>"> </td>
            <td colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Item 
              &nbsp;Desabilitado por seguran&ccedil;a</font></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td bgcolor="#EEEEEE"><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Cor 
                Links:</font></div></td>
            <td><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <input name="cor_links" type="text"  class="text" id="cor_links" style="width:195px;" value="<%= Rs("cor_links")%>" >
              </font><font size="1" face="Arial, Helvetica, sans-serif"><strong><font color="#006600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../imagens/forum_images/post_button_colour_pallete.gif" width="25" height="24" align="absmiddle" border="0" alt="Escolher cores" onClick="foreColor();" style="cursor: hand;"></font></strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              </font></strong></td>
            <td><div align="right"><font color="#FF0000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ex:</font></div></td>
            <td><div style="color: <%= Rs("cor_links")%>; " id="cor_link"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Link</font></div></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td bgcolor="#EEEEEE"><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Cor 
                fundo Box:</font></div></td>
            <td><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <input name="cor_fundo_box" type="text"  class="text" id="cor_fundo_box" style="width:195px;" value="<%=cor_fundo_box%>" >
              </font><font size="1" face="Arial, Helvetica, sans-serif"><strong><font color="#006600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../imagens/forum_images/post_button_fill.gif" width="25" height="24" align="absmiddle" 
			  border="0" alt="Escolher cores" onClick="foreColor2();" style="cursor: hand;"></font></strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              </font></strong></td>
            <td><div align="right"><font color="#FF0000" face="Verdana, Arial, Helvetica, sans-serif"><font size="1">Ex:</font></font></div></td>
            <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <input name="ex_cor_fundo_box" type="text"   Class="text" style="background-color: <%=cor_fundo_box%>; " value="Texto" size="10"  >
              </font></strong></font></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td bgcolor="#EEEEEE"><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Fonte 
                Site:</font></div></td>
            <td><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <input name="site_fonte" type="text"  class="text" id="site_fonte" style="width:130px;" value="<%= Rs("site_fonte")%>" >
              <select name="fonte" 
						 class="text"
						  onChange="fontesete(fonte.options[fonte.selectedIndex].value)" >
                <option value="0">Fontes</option>
                <option value="Arial, Helvetica, sans-serif">Arial</option>
                <option value="Times New Roman, Times, serif">Times</option>
                <option value="Courier New, Courier, mono">Courier New</option>
                <option value="Verdana, Arial, Helvetica, sans-serif">Verdana</option>
              </select>
              </font></strong></td>
            <td><div align="right"><font color="#FF0000" face="Verdana, Arial, Helvetica, sans-serif"><font size="1">Ex:</font></font></div></td>
            <td><div style="font-family: <%= Rs("site_fonte")%>; " id="tipo_fonte2">Tipo 
                Fonte</div></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td bgcolor="#EEEEEE"><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Cor 
                fundo Topo Site:</font></div></td>
            <td> <input name="cor_fundo_site" type="text"  class="text"  style="width:195px;" value="<%=cor_fundo_site%>" > 
              <img src="../imagens/forum_images/post_button_fill.gif" width="25" height="24" align="absmiddle" 
			  border="0" alt="Escolher cores" onClick="foreColor3();" style="cursor: hand;"> </font></strong> 
            </td>
            <td><div align="right"><font color="#FF0000" face="Verdana, Arial, Helvetica, sans-serif"><font size="1">Ex:</font></font></div></td>
            <td bgcolor="#EEEEEE"> <div style="background-color: <%=cor_fundo_site%>;   padding: 1; height: 21px; width: 50px;border-style: solid; border-color: #666666;border-size:1; border-color: #666666;" id="m_cor_fundo_site" ></div></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td bgcolor="#EEEEEE">&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td bgcolor="#EEEEEE"><strong><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></strong></td>
            <td colspan="3"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
              </font></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td bgcolor="#EEEEEE"> <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"> 
                </font></div></td>
            <td bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
              </font></td>
            <td bgcolor="#EEEEEE">&nbsp;</td>
            <td bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif"> 
              <input name="imageField" type="image" src="../imagens/outros_icos/botao_atualizar_dados.gif" width="129" height="24" border="0">
              </font></td>
          </tr>
        </table>
  
  
  </td>
  </tr>
</table>
	 <!-- fIM tab -->
  
  
	  </form>







	<%
	Rs.Close
   	set Rs = nothing
	Conn.Close
	set Conn = nothing



if request.QueryString("config_site") = "atualizar" Then

		Set Conn = Server.CreateObject("ADODB.Connection")
   		Conn.Open base_dados 
   	      		   
	       Sql="UPDATE  configuracoes_site SET site_titulo = '" & Request.form("site_titulo") &  "', site_end = '" & Request.form("site_end") &  "',"
		   sql = sql + "host_mail = '" & Request.form("host_mail") &  "', cor_fundo_box = '" & Request.form("cor_fundo_box") &  "',"
		   sql = sql + " admin_mail = '" & Request.form("admin_mail") &  "', logo_empresa = '" & Request.form("logo_empresa") &  "',"
		 	sql = sql + " dados = '" & Request.form("dados") &  "', cor_links = '" & Request.form("cor_links") &  "', cor_fundo_site = '" & Request.form("cor_fundo_site") &  "',"
			sql = sql + " site_fonte = '" & Request.form("site_fonte") &  "' WHERE id_geral =1"       		
	       Conn.Execute(Sql)
	   %>
	   <script>
			msg = "   DADOS ATUALIZADO COM SUCESSO     \n";
		msg += "_____________________________________\n\n";
		msg += "Os dados forama tualizados na tabela com sucesso!:.\n";
		msg += "=================================================================\n";
		
		msg += "A equipe do <%=site_titulo%> agradece!!!! Você está sendo redirecionado para nossa página inicial. \n";
		
		alert(msg + "\n\n");
		location.href="config_site_admin.asp"
		
		</script>
	   <%
	   
End If

	%>
	
<!--#include file="rodape.asp"-->

