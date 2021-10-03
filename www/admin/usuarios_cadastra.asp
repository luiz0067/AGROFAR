<!-- #include file="login_nivel3.asp" -->

<script language="JavaScript">

  function verify(form){
        if(form.nome_promo.value == ""){
		
		
		msg = "           CAMPO OBRIGATÓRIO         \n";
		msg += "_____________________________________\n\n";
		msg += "Você não preencheu o campo Nome da Promoção.\n";
		msg += "Este é uma campo Obrigatório\n\n";
		msg += "Você está sendo direcionado para o campo em branco a ser preenchido \n";
		
		alert(msg + "\n\n");
		
            form.nome_promo.focus();
            return false;
      }
      if(form.local.value == ""){
	  
	  msg = "           CAMPO OBRIGATÓRIO         \n";
		msg += "_____________________________________\n\n";
		msg += "Você não preencheu o campo Local, cidade onde vaiser a promoção.\n";
		msg += "Este é uma campo Obrigatório\n\n";
		msg += "Você está sendo direcionado para o campo em branco a ser preenchido \n";
		
		alert(msg + "\n\n");

            form.local.focus();
            return false;
      }
      if(form.data_ini.value == ""){
	  
	  	msg = "           CAMPO OBRIGATÓRIO         \n";
		msg += "_____________________________________\n\n";
		msg += "Você não preencheu o campo com a Data do começo da promoção.\n";
		msg += "Este é uma campo Obrigatório\n\n";
		msg += "Você está sendo direcionado para o campo em branco a ser preenchido \n";
		
		alert(msg + "\n\n");

            form.data_ini.focus();
            return false;
      }
      
      if(form.data_fim.value == ""){
	  
	  			msg = "           CAMPO OBRIGATÓRIO         \n";
		msg += "_____________________________________\n\n";
		msg += "Você não preencheu o campo com a Data do Fim da Promoção.\n";
		msg += "Este é uma campo Obrigatório\n\n";
		msg += "Você está sendo direcionado para o campo em branco a ser preenchido \n";
		
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


<form  method="post" action="usuarios_cadastra.asp?usuarios=new" onSubmit="return verify(this)" name="config">
  <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><br>
    Passe o mouse sobre <strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../imagens/outros_icos/sm-unknown1.gif" width="20" height="20"></font></strong>ícone 
    para obter informações sobre o campo.<br>
    <font color="#FF0000" size="3"><em><dd>*</dd></em></font><font color="#FF0000"> todos 
    os campos são obrigatórios. </font></font><br><br>
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
              <input name="nome" type="text" id="nome"  class="text" style="width:195px;" > 
            </td>
          </tr>
          <tr> 
            <td height="24" bgcolor="#EEEEEE"> <div align="right" background="../imagens/forum_images/table_bg2.gif"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Sobrenome:</font></strong></div></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
              <input name="sobrenome" type="text" id="sobrenome"  class="text" style="width:195px;">
              </font></td>
          </tr>
          <tr> 
            <td height="24" bgcolor="#EEEEEE"><div align="right"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">E-mail:</font></strong></div></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
              <input name="email" type="text" id="email"  class="text" style="width:195px;">
              </font></td>
          </tr>
          <tr> 
            <td height="24" bgcolor="#EEEEEE"><div align="right"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Login:</font></strong></div></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
              <input name="login" type="text" id="login"  class="text" style="width:195px;">
              </font></td>
          </tr>
          <tr>
            <td height="24" bgcolor="#EEEEEE"><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Senha:</strong></font></div></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
              <input name="senha" type="text" id="senha"  class="text" style="width:195px;">
              </font></td>
          </tr>
          <tr> 
            <td height="24" bgcolor="#EEEEEE"><div align="right"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Endere&ccedil;o</font></strong></div></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
              <input name="endereco" type="text" id="endereco"  class="text" style="width:195px;">
              </font></td>
          </tr>
          <tr> 
            <td height="24" bgcolor="#EEEEEE"><div align="right"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade:</font></strong></div></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
              <input name="cidade" type="text" id="cidade"  class="text" style="width:110px;">
              <font face="Verdana, Arial, Helvetica, sans-serif"><strong>Estado:</strong></font> 
              <select name="estado"  class="text" style="width:40px;">
                <OPTION value=null 
                                selected>UF:</OPTION>
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
              <select name="dia" id="dia" class="text" style="width:40px;">
                <option>Dia</option>
                <option value="1">1</option>
                <option value="2">2</option>
                <option value="3">3</option>
                <option value="4">4</option>
                <option value="5">5</option>
                <option value="6">6</option>
                <option value="7">7</option>
                <option value="8">8</option>
                <option value="9">9</option>
                <option value="10">10</option>
                <option value="11">11</option>
                <option value="12">12</option>
                <option value="13">13</option>
                <option value="14">14</option>
                <option value="15">15</option>
                <option value="16">16</option>
                <option value="17">17</option>
                <option value="18">18</option>
                <option value="19">19</option>
                <option value="20">20</option>
                <option value="21">21</option>
                <option value="22">22</option>
                <option value="23">23</option>
                <option value="24">24</option>
                <option value="25">25</option>
                <option value="26">26</option>
                <option value="27">27</option>
                <option value="28">28</option>
                <option value="29">29</option>
                <option value="30">30</option>
                <option value="31">31</option>
              </select>
              <font color="#000000" size="2">/</font> 
              <select name="mes" id="mes" class="text" style="width:40px;">
                <option>M&ecirc;s</option>
                <option value="1">1</option>
                <option value="2">2</option>
                <option value="3">3</option>
                <option value="4">4</option>
                <option value="5">5</option>
                <option value="6">6</option>
                <option value="7">7</option>
                <option value="8">8</option>
                <option value="9">9</option>
                <option value="10">10</option>
                <option value="11">11</option>
                <option value="12">12</option>
              </select>
              <font color="#000000">19</font> </font></strong> 
              <input name="ano" type="text"  class="text" id="ano" style="width:30px;" maxlength="2">
              </font></td>
          </tr>
          <tr> 
            <td height="24" bgcolor="#EEEEEE"><div align="right"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Sexo:</font></strong></div></td>
            <td colspan="2" bgcolor="#EEEEEE"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <input name="sexo" type="radio" value="M" checked>
              Masculino 
              <input type="radio" name="sexo" value="F">
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
                <option>-------- Escolha um N&iacute;vel --------</option>
                <option value="1">:: Acesso Normal = 1 :: </option>
                <option value="2">:: Acesso Limitado = 2 :: </option>
                <option value="3">:: Acesso Administrador = 3 :: </option>
              </select> <strong></strong></td>
            <td width="29%" bgcolor="#EEEEEE"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../imagens/outros_icos/sm-unknown1.gif" width="20" height="20"
			style="cursor:help;" onmouseout="kill()" onmouseover="popup('<div align=left><font color=#ffffff><b> Cada usuário possui um nívelde acesso determinando o que pode e o que não pode fazer.</b><hr><b>Nível Normal- </b>ler<br><b>Nível Limitado - </b>pode deletar e editar somente em algumas sessões<br><b>Administrador - </b>tem permissão total no sistema.','#6699CC')"></font></strong></td>
          </tr>
          <tr> 
            <td height="24" bgcolor="#EEEEEE"><div align="right"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Alerta 
                novas Mensagens:</font></strong></div></td>
            <td bgcolor="#EEEEEE">&nbsp;<strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <input name="alert_msg" type="radio" value="True" checked>
              Sim 
              <input type="radio" name="alert_msg" value="False">
              N&atilde;o</font></strong></td>
            <td bgcolor="#EEEEEE"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../imagens/outros_icos/sm-unknown1.gif" width="20" height="20"
			style="cursor:help;" onmouseout="kill()" onmouseover="popup('<div align=left><font color=#ffffff><b> Quando você recebe uma E-mail (mensagem) do sistema, vc automaticamente recebe um alerta bem no topo da página, e outro que é uma mensagem que aparece na sua tela, perguntando se você deseja ler suas mensagen(s). Aqui você pode setar se a pessao recbe ou não este alerta.','#6699CC')"></font></strong></td>
          </tr>
          <tr> 
            <td height="24" bgcolor="#EEEEEE"><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Iacute;cone/Foto 
                Usu&aacute;rio:</strong></font>:</div></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
              <select name="SelectAvatar" size="4" class="text" style="width:195px;"
		  onChange="(avatar.src = SelectAvatar.options[SelectAvatar.selectedIndex].value) && (foto.value=SelectAvatar.options[SelectAvatar.selectedIndex].value)">
                <option value="../imagens/avatars/blank.gif" >Sem Imagem</option>
                <!-- #include file="../functions/select_avatar.asp" -->
              </select>
              <img src="<%

		'If there is an avatar then display it
		If strAvatar <> "" Then
		     	Response.Write(strAvatar)
		Else
			Response.Write( "../imagens/avatars/blank.gif")
		End If
                %>" width="64" height="64" name="avatar"> 
              <input type="hidden" name="foto" value="<% = strAvatar %>">
              </font></td>
          </tr>
          <tr> 
            <td height="24" bgcolor="#EEEEEE"><div align="right"></div></td>
            <td colspan="2" bgcolor="#EEEEEE"> <input type="button"   
			  onClick="window.open('usuarios_upload.asp','avatars','toolbar=0,location=0,status=1,menubar=0,scrollbars=0,resizable=1,width=400,height=150')"
			  value="Clique aqui para Enviar a sua Foto/Ícone" class="text" style="width:205px;cursor:hand;"></td>
          </tr>
          <tr> 
            <td bgcolor="#EEEEEE"><strong><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></strong></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
              </font></td>
          </tr>
          <tr> 
            <td bgcolor="#EEEEEE"> <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"> 
                </font></div></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif"> 
              <input name="cad" type="image" src="../imagens/outros_icos/botao_cadastrar_acao.gif" width="129" height="24" border="0">
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

if Request.QueryString("usuarios") = "new" Then 
	 Set Conn = Server.CreateObject("ADODB.Connection")
 	Conn.Open base_dados
	Set news = Server.CreateObject("ADODB.Recordset")
   		Sql = "SELECT * FROM usuarios WHERE login='" & Request.Form("login") & "'"
	   news.Open Sql, base_dados
	   'Conn.Execute(Sql)
	   IF Not news.EOF THEN
	   %>
	   <script>
	   
	   	msg = "           LOGIN JÁ EXISTTENTE         \n";
		msg += "_____________________________________\n\n";
		msg += "O login já existe em nosso banco de dados.\n";
		msg += "Por favor escolha outro Login\n\n";
		msg += "Você está sendo direcionado para o campo para altera-lo\n";
		
		alert(msg + "\n\n");
		history.go(-1);
        form.login.focus(-1);

	   
	   </script>
	   <%
	   Else
	      Sql = "INSERT INTO usuarios (nome,sobrenome,email,login,senha,endereco,cidade,estado,sexo,niver,data,acesso,newsletter,alert_msg,foto)"
		  sql = sql + " VALUES('" & Request.Form("nome") & "','" & Request.Form("sobrenome") & "','" & Request.Form("email") & "',"
		  sql = sql + "'" & Request.Form("login") & "','" & Request.Form("senha") & "','" & Request.Form("endereco") & "',"
		  sql = sql + "'" & Request.Form("cidade") & "','" & Request.Form("estado") & "','" & Request.Form("sexo") & "',"
		  sql = sql + "'" & Request.Form("dia") & ("/") & Request.Form("mes") & ("/") & Request.Form("ano") &"',"
		  sql = sql + "'" & Day(Now) & ("/") & month(Now) & ("/") & Year(Now) & "','" & Request.Form("acesso") & "',TRUE," & Request.Form("alert_msg") & ",'" & Request.Form("foto") & "')"
		  Conn.Execute(Sql)
		  		
   
   %><script>
	   
	   	msg = "           USUÁRIO CADASTRADO COM SUCESSO!!!         \n";
		msg += "_____________________________________________________\n\n";
		msg += "Seus dados foram cadastrados com Sucesso no nosso Ba †††† ????l???nco de Dados.\n\n";
		msg += "Agora você passa a usufruir de vários serviços do nosso site.\n";
		msg += "Para alterar seus dados basta indentificar-se no formulário ao lado,\n";
		msg += "e ir na sua configurações e usuário.\n\n";
		msg += "Obrigado por fazer parte do time de usuários da Dominiun Montanhismo.\n";
		
		alert(msg + "\n\n");
		location.href="usuarios_default_admin.asp";

	   
	   </script>
<%
   			
   			end if
   end if
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
Yoffset= 10;    // mudar a posição do popup.

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

