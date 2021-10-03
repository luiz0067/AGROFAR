<!-- #include file="login_nivel3.asp" -->
<%
		' Abre o banco de dados e seleciona o id que quer ser lido
		set Rs = server.createobject("ADODB.recordset")
		Sqltext = "SELECT * FROM configuracoes_site"
		Set Rs = Conn.Execute(Sqltext)
%>

<FORM ACTION="?config_site=atualizar" METHOD="POST" name="config" id="config" >
  
  <table width="400" border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center">
  <tr>
    <td>
	
	
  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
          <tr > 
            <td height="23" colspan="5" background="../imagens/forum_images/table_bg.gif"> 
              <center>
                <strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">:: 
                Telefones Import&acirc;ntes da Empresa::</font> </strong> </center></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td width="31%" bgcolor="#EEEEEE"> <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;T&iacute;tulo 
                Topo Fone <font color="#FF0000">1</font>: </font></div></td>
            <td colspan="2"> <font size="1" face="Arial, Helvetica, sans-serif"> 
              <input name="fone_topo_tit_1" type="text" id="fone_topo_tit_1"  class="text" style="width:195px;" value="<%= Rs("fone_topo_tit_1")%>" >
              </font> </td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td width="31%" height="24" bgcolor="#EEEEEE"> <div align="right" background="../imagens/forum_images/table_bg2.gif"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Fone 
                Topo <font color="#FF0000">1</font>:</font></div></td>
            <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif"> 
              <input name="fone_topo_1" type="text"  class="text" id="fone_topo_1" style="width:195px;" value="<%= Rs("fone_topo_1")%>">
              </font></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td bgcolor="#EEEEEE" width="35%"></td>
            <td bgcolor="#cococo" width="55%"></td>
            <td bgcolor="#EEEEEE" width="10%"></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td width="31%" bgcolor="#EEEEEE"><div align="right"><font color="#000000" size="2">&nbsp;</font><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">T&iacute;tulo 
                Topo Fone <font color="#FF0000">2</font>:</font></div></td>
            <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif"> 
              <input name="fone_topo_tit_2" type="text" id="fone_topo_tit_2"  class="text" style="width:195px;" value="<%= Rs("fone_topo_tit_2")%>" >
              </font></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td width="31%" bgcolor="#EEEEEE"><div align="right"><font color="#000000" size="2">&nbsp;</font><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Fone 
                Topo <font color="#FF0000">2</font>:</font></div></td>
            <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif"> 
              <input name="fone_topo_2" type="text"  class="text" id="fone_topo_2" style="width:195px;" value="<%= Rs("fone_topo_2")%>">
              </font></td>
          </tr> <tr bgcolor="#EEEEEE"> 
            <td bgcolor="#EEEEEE" width="35%"></td>
            <td bgcolor="#cococo" width="55%"></td>
            <td bgcolor="#EEEEEE" width="10%"></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td bgcolor="#EEEEEE"><div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">T&iacute;tulo 
                Topo Fone <font color="#FF0000">3</font>:</font></div></td>
            <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif"> 
              <input name="fone_topo_tit_3" type="text" id="fone_topo_tit_3"  class="text" style="width:195px;" value="<%= Rs("fone_topo_tit_3")%>" >
              </font></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td bgcolor="#EEEEEE"><div align="right"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Fone 
                Topo <font color="#FF0000">3</font>:</font></div></td>
            <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif"> 
              <input name="fone_topo_3" type="text"  class="text" id="fone_topo_122" style="width:195px;" value="<%= Rs("fone_topo_3")%>">
              </font></td>
          </tr> <tr bgcolor="#EEEEEE"> 
            <td bgcolor="#EEEEEE" width="35%"></td>
            <td bgcolor="#cococo" width="55%"></td>
            <td bgcolor="#EEEEEE" width="10%"></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td bgcolor="#EEEEEE"> <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"> 
                </font></div></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
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
   	      		   
	       Sql="UPDATE  configuracoes_site SET fone_topo_tit_1 = '" & Request.form("fone_topo_tit_1") &  "', fone_topo_tit_2 = '" & Request.form("fone_topo_tit_2") &  "',"
		   sql = sql + "fone_topo_tit_3 = '" & Request.form("fone_topo_tit_3") &  "', fone_topo_1 = '" & Request.form("fone_topo_1") &  "',"
		   sql = sql + " fone_topo_2 = '" & Request.form("fone_topo_2") &  "', fone_topo_3 = '" & Request.form("fone_topo_3") &  "' "
		 	sql = sql + " WHERE id_geral =1"       		
	       Conn.Execute(Sql)
	   %>
	   <script>
			msg = "   DADOS ATUALIZADO COM SUCESSO     \n";
		msg += "_____________________________________\n\n";
		msg += "Os dados forama tualizados na tabela com sucesso!:.\n";
		msg += "=================================================================\n";
		
		msg += "A equipe do <%=site_titulo%> agradece!!!! Você está sendo redirecionado para nossa página inicial. \n";
		
		alert(msg + "\n\n");
		location.href="fones_topo.asp"
		
		</script>
	   <%
	   
End If

	%>
	
<!--#include file="rodape.asp"-->

