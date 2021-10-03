<!-- #include file="login_nivel3.asp" -->


<script language="JavaScript">

  function verify(config){
        if(config.nome.value == ""){
		
		
		msg = "           CAMPO OBRIGATÓRIO         \n";
		msg += "_____________________________________\n\n";
		msg += "Voce nao preencheu o campo Nome.\n";
		msg += "Este é uma campo Obrigatório\n\n";
		msg += "Voce está sendo direcionado para o campo em branco a ser preenchido \n";
		
		alert(msg + "\n\n");
		
            form.nome.focus();
            return false;
      }
      if(form.sobrenome.value == ""){
	  
	  msg = "           CAMPO OBRIGATÓRIO         \n";
		msg += "_____________________________________\n\n";
		msg += "Voce nao preencheu o campo Sobrenome.\n";
		msg += "Este é uma campo Obrigatório\n\n";
		msg += "Voce está sendo direcionado para o campo em branco a ser preenchido \n";
		
		alert(msg + "\n\n");

            form.sobrenome.focus();
            return false;
      }
      if(form.data_ini.value == ""){
	  
	  	msg = "           CAMPO OBRIGATÓRIO         \n";
		msg += "_____________________________________\n\n";
		msg += "Voce nao preencheu o campo com a Data do começo da promoçao.\n";
		msg += "Este é uma campo Obrigatório\n\n";
		msg += "Voce está sendo direcionado para o campo em branco a ser preenchido \n";
		
		alert(msg + "\n\n");

            form.data_ini.focus();
            return false;
      }
      
      if(form.data_fim.value == ""){
	  
	  			msg = "           CAMPO OBRIGATÓRIO         \n";
		msg += "_____________________________________\n\n";
		msg += "Voce nao preencheu o campo com a Data do Fim da Promoçao.\n";
		msg += "Este é uma campo Obrigatório\n\n";
		msg += "Voce está sendo direcionado para o campo em branco a ser preenchido \n";
		
		alert(msg + "\n\n");

            form.data_fim.focus();
            return false;
            
      }
     

  }


function foreColor()	{
	var arr = showModalDialog("calendario.asp",window,"font-family:Verdana; font-size:12; status:false; dialogWidth:14; dialogHeight:14" );
	if (arr != null) cmdExec("ForeColor",arr);	
}
function cmdExec(cmd,opt) {
	document.config.data_ini.value=(cmd,"","" + opt);
	config.data_ini.focus();
}
function foreColor2()	{
	var arr = showModalDialog("calendario.asp",window,"font-family:Verdana; font-size:12; status:false; dialogWidth:14; dialogHeight:14" );
	if (arr != null) cmdExec2("ForeColor",arr);	
}
function cmdExec2(cmd,opt) {
	document.config.data_fim.value=(cmd,"","" + opt);
	config.data_fim.focus();
}
function fontesete(opt) {
	document.config.site_fonte.value=(opt);
	config.site_fonte.focus();
}
</script>
<%
Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open base_dados
	
	Set Rs = Server.CreateObject("ADODB.Recordset")
   Sql="SELECT * from usuarios WHERE id =" & Request.QueryString("id")
   Rs.Open Sql, base_dados
   %>

<form  method="post" action="usuarios_editar.asp?usuarios=atualizar&id=<%=request("id")%>" onSubmit="return verify(this)" name="config">
  <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
    Passe o mouse sobre <strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../imagens/outros_icos/sm-unknown1.gif" width="20" height="20"></font></strong>ícone 
    para obter informaçoes sobre o campo.<br>
    <font color="#FF0000" size="3"><em><dd>*</dd></em></font><font color="#FF0000"> todos 
    os campos sao obrigatórios. </font></font>
    <!-- Inicio tab -->
  </div>
  <table width="41%" border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center">
  <tr>
    <td>
	
	
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr  > 
            <td height="23" colspan="3" background="../imagens/forum_images/table_bg.gif"> 
              <center>
                <strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">:: 
                Cadastrar novo Usu&aacute;rio no Sistema::</font> </strong> </center></td>
          </tr>
          <tr> 
            <td width="30%" bgcolor="#EEEEEE"> <div align="right"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;Nome:</font></strong></div></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font> 
              <input name="nome" type="text"  class="text" id="nome" style="width:195px;" value="<%=Rs("nome")%>" >            </td>
          </tr>
          <tr> 
            <td height="24" bgcolor="#EEEEEE"><div align="right"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">E-mail:</font></strong></div></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
              <input name="email" type="text"  class="text" id="email" style="width:195px;" value="<%=Rs("email")%>">
              </font></td>
          </tr>
          <tr> 
            <td height="24" bgcolor="#EEEEEE"><div align="right"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Login:</font></strong></div></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
              <input name="login" type="text"  class="text" id="login" style="width:195px;" value="<%=Rs("login")%>">
              </font></td>
          </tr>
          <tr> 
            <td height="24" bgcolor="#EEEEEE"><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Senha:</strong></font></div></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
              <input name="senha" type="password"  class="text" id="senha" style="width:195px;" value="<%=Rs("senha")%>">
              </font></td>
          </tr>
          <tr> 
            <td height="24" bgcolor="#EEEEEE"><div align="right"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Endere&ccedil;o</font></strong></div></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
              <input name="endereco" type="text"  class="text" id="endereco" style="width:195px;" value="<%=Rs("endereco")%>">
              </font></td>
          </tr>
          <tr> 
            <td height="24" bgcolor="#EEEEEE"><div align="right"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade:</font></strong></div></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
              <input name="cidade" type="text"  class="text" id="cidade" style="width:110px;" value="<%=Rs("cidade")%>">
              <font face="Verdana, Arial, Helvetica, sans-serif"><strong>Estado:</strong></font> 
              <select name="estado"  class="text" style="width:40px;">
                <OPTION value=<%=Rs("estado")%> 
                                selected><%=Rs("estado")%></OPTION>
                <OPTION 
                                value=null>--</OPTION>
                <OPTION 
                                value=AC>AC</OPTION>
                <OPTION 
                                value=AL>AL</OPTION>
                <OPTION 
                                value=AM>AM</OPTION>
                <OPTION 
                                value=AP>AP</OPTION>
                <OPTION 
                                value=BA>BA</OPTION>
                <OPTION 
                                value=CE>CE</OPTION>
                <OPTION 
                                value=DF>DF</OPTION>
                <OPTION 
                                value=ES>ES</OPTION>
                <OPTION 
                                value=GO>GO</OPTION>
                <OPTION 
                                value=MA>MA</OPTION>
                <OPTION 
                                value=MG>MG</OPTION>
                <OPTION 
                                value=MS>MS</OPTION>
                <OPTION 
                                value=MT>MT</OPTION>
                <OPTION 
                                value=PA>PA</OPTION>
                <OPTION 
                                value=PB>PB</OPTION>
                <OPTION 
                                value=PE>PE</OPTION>
                <OPTION 
                                value=PI>PI</OPTION>
                <OPTION 
                                value=PR>PR</OPTION>
                <OPTION 
                                value=RJ>RJ</OPTION>
                <OPTION 
                                value=RN>RN</OPTION>
                <OPTION 
                                value=RO>RO</OPTION>
                <OPTION 
                                value=RR>RR</OPTION>
                <OPTION 
                                value=RS>RS</OPTION>
                <OPTION 
                                value=SC>SC</OPTION>
                <OPTION 
                                value=SE>SE</OPTION>
                <OPTION 
                                value=SP>SP</OPTION>
                <OPTION 
                                value=TO>TO</OPTION>
              </select>
              </font></td>
          </tr>
          <tr> 
            <td height="24" bgcolor="#EEEEEE"><div align="right"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
                Nascimento:</font></strong></div></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
              <strong><font color="#006600" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              </font></strong> 
              <input name="niver" type="text"  class="text"  style="width:110px;" value="<%=Rs("niver")%>" maxlength="12">
              <font face="Verdana, Arial, Helvetica, sans-serif"><strong>dd/mm/aaaa 
              </strong></font></font></td>
          </tr>
          <tr> 
            <td height="24" bgcolor="#EEEEEE"><div align="right"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Sexo:</font></strong></div></td>
            <td colspan="2" bgcolor="#EEEEEE"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <% If Rs("sexo") = "M" Then 	%>
              <input name="sexo" type="radio" value="M" checked>
              Masculino 
              <input type="radio" name="sexo" value="F">
              <%	Else	%>
              <input name="sexo" type="radio" value="M" >
              Masculino 
              <input type="radio" name="sexo" value="F" checked>
              <%	End If	%>
              Feminino</font></strong></td>
          </tr>
          <tr > 
            <td height="24" colspan="3" background="../imagens/forum_images/table_bg.gif"> 
              <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>:: 
                Configura&ccedil;&otilde;es do Usu&aacute;rio ::</strong></font></div></td>
          </tr>
          <tr> 
            <td height="24" bgcolor="#EEEEEE"><div align="right"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">N&iacute;vel 
                de Acesso:</font></strong></div></td>
            <td width="41%" bgcolor="#EEEEEE">&nbsp; <select size="1" name="acesso" class="text" style="width:195px;">
                <OPTION value=<%=Rs("acesso")%> selected><%=Rs("acesso")%></OPTION>
                <option>-------- Escolha um N&iacute;vel --------</option>
                <option value="1">:: Acesso Normal = 1 :: </option>
                <option value="2">:: Acesso Limitado = 2 :: </option>
                <option value="3">:: Acesso Administrador = 3 :: </option>
              </select> <strong></strong></td>
            <td width="29%" bgcolor="#EEEEEE"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../imagens/outros_icos/sm-unknown1.gif" width="20" height="20"
			style="cursor:help;" onMouseOut="kill()" onMouseOver="popup('<div align=left><font color=#ffffff><b> Cada usuário possui um nívelde acesso determinando o que pode e o que nao pode fazer.</b><hr><b>Nível Normal- </b>ler<br><b>Nível Limitado - </b>pode deletar e editar somente em algumas sessoes<br><b>Administrador - </b>tem permissao total no sistema.','#6699CC')"></font></strong></td>
          </tr>
          <tr> 
            <td height="24" bgcolor="#EEEEEE"><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Iacute;cone/Foto 
                Usu&aacute;rio:</strong></font>:</div></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
              <select name="SelectAvatar" size="4" class="text" style="width:195px;"
		  onChange="(avatar.src = SelectAvatar.options[SelectAvatar.selectedIndex].value) && (foto.value=SelectAvatar.options[SelectAvatar.selectedIndex].value)">
                <option value="../imagens/avatars/blank.gif">Sem Imagem</option>
                <!-- #include file="select_avatar.asp" -->
              </select>
              <img src="<%

		'If there is an avatar then display it
		If rs("foto") <> "" Then
		     	Response.Write(Rs("foto"))
		Else
			Response.Write( "../imagens/avatars/blank.gif")
		End If
                %>" width="64" height="64" name="avatar"> 
              <input type="hidden" name="foto" value="<%=Rs("foto")%>">
              </font></td>
          </tr>
          <tr> 
            <td height="24" bgcolor="#EEEEEE"><div align="right"></div></td>
            <td colspan="2" bgcolor="#EEEEEE"> <input type="button"   
			  onClick="window.open('usuarios_upload.asp','avatars','toolbar=0,location=0,status=1,menubar=0,scrollbars=0,resizable=1,width=400,height=150')"
			  value="Clique aqui para Enviar a sua Foto/Ícone" class="text" style="width:205px;cursor:hand;font-size:9pt"></td>
          </tr>
          <tr> 
            <td bgcolor="#EEEEEE"> <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"> 
                </font></div></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif"> 
              <input name="cad" type="image" src="../imagens/outros_icos/botao_atualizar_dados.gif" width="129" height="24" border="0">
              </font></td>
          </tr>
          <tr> 
            <td colspan="3" bgcolor="#EEEEEE">&nbsp;</td>
          </tr>
        </table>
  
  
  </td>
  </tr>
</table>

	 <!-- fIM tab -->
</form>
<%

' --------------------------------------------------

Conn.Close
   Set Conn = Nothing
   Rs.Close
   Set Rs = Nothing
   
   ' ************************************************************************************
   ' Atualizar usuários
   ' ************************************************************************************
   
   If Request.QueryString("usuarios") = "atualizar" Then
    Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open base_dados 
	
   	if Request("newsletter") = "TRUE" Then
		newsletter = "TRUE"
		Else
		newsletter = "FALSE"
	ENd If

   	      		   
	       Sql="UPDATE  usuarios SET nome = '" & Request("nome") &  "', cidade = '" & Request("cidade") &  "', foto = '" & Request("foto") &  "', "
		   sql = sql + "estado = '" & Request("estado") &  "', endereco = '" & Request("endereco") &  "',"
		   sql = sql + "email = '" & Request("email") & "', sexo = '" & Request("sexo") & "', "
		   sql = sql + "senha = '" & Request("senha") & "', login = '" & Request("login") & "',niver = '" & Request("niver") & "',acesso = '" & Request("acesso") & "',"
		   sql = sql + "newsletter = " & newsletter & " WHERE id =" & request("id")     		
	       Conn.Execute(Sql)
  		
   nome = request("nome")
   %>
   <script>
	   
	   	msg = "           USUÁRIO ATUALIZADO COM SUCESSO!!!         \n";
		msg += "_____________________________________________________\n\n";
		msg += "Os dados de <%=nome%> foram Atualizados com Sucesso no nosso Banco de Dados.\n\n";
		
		alert(msg + "\n\n");
		location.href="usuarios_default_admin.asp";

	   
	   </script>
<%
   
   			
			Conn.Close
			Set Conn = Nothing
   'news.Close
   'Set news = nothing
			
End If
      %>
<!--#include file="rodape.asp"-->
<!-- MANTER ESSE CÓDIGO NO FINAL DO SEU CÓDIGO HTML, ANTES DO FINAL DO TAG BODY //-->
<DIV ID="dek" CLASS="dek"></DIV>
<STYLE TYPE="text/css">
<!--
BODY {OVERFLOW:scroll;OVERFLOW-X:hidden}
.DEK {POSITION:absolute;VISIBILITY:hidden;Z-INDEX:200;}
//-->
</STYLE>
<SCRIPT TYPE="text/javascript">
<!--
Xoffset=20;    // modifique esses valores para ...
Yoffset= 10;    // mudar a posiçao do popup.

var nav,old,iex=(document.all),yyy=-1000;
if(navigator.appName=="Netscape"){(document.layers)?nav=true:old=true;}

if(!old){
var skn=(nav)?document.dek:dek.style;
if(nav)document.captureEvents(Event.MOUSEMOVE);
document.onmousemove=get_mouse;
}

function popup(msg,bak){
var content="<table border=0 width=200 height=25 cellpading=1 cellspacing=1 "+
"bgcolor="+bak+" style='filter:Alpha(Opacity=60,FinishOpacity=90,Style=2,StartX=100,StartY=100,FinishX=100,FinishY=1);font-family:verdana;font-size:9pt;color:#ffffff;font-weight:;'><td align=center valign=middle>"+msg+"</td></tr></table>";
if(old){alert(msg);return;} 
else{yyy=Yoffset;
 if(nav){skn.document.write(content);skn.document.close();skn.visibility="visible"}
 if(iex){document.all("dek").innerHTML=content;skn.visibility="visible"}
 }
}

function get_mouse(e){
var x=(nav)?e.pageX:event.x+document.body.scrollLeft;skn.left=x+Xoffset;
var y=(nav)?e.pageY:event.y+document.body.scrollTop;skn.top=y+yyy;
}

function kill(){
if(!old){yyy=-1000;skn.visibility="hidden";}
}
//-->
</SCRIPT>
<!-- MANTER ESSE CÓDIGO NO FINAL DO SEU CÓDIGO HTML, ANTES DO FINAL DO TAG BODY //-->

