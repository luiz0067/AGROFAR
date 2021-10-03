<!-- #include file="login_nivel3_limpo.asp" -->

<title>:: Mostrando Detalhes do Prospect:: </title>
<body background="../imagens/forum_images/bg.gif">
<br>
<%
Set Rs = Server.CreateObject("ADODB.Recordset")
pro= "SELECT * FROM usuarios where id= " & request("id") 
Rs.Open pro, base_dados
%>
<center>
<!-- Inicio tab -->
            <table width="650" border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center">
  <tr>
    <td>
	
	
  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
          <tr > 
            <td height="23" colspan="5" background="../imagens/forum_images/table_bg.gif"> 
              <center>
                <strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">:: 
                Mostrando Detalhes do Usu&aacute;rio::</font> </strong> </center></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td width="13%" rowspan="4" bgcolor="#EEEEEE"><img src="../<%=Rs("foto")%>" width="64" height="64"></td>
            <td width="20%" bgcolor="#EEEEEE"> <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>&nbsp;Nome: 
                </strong></font></div></td>
            <td width="16%"> <strong><font color="#FF0000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%=Rs("nome")%> <%=Rs("sobrenome")%></font></strong></td>
            <td width="16%"> <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Login 
                :</strong></font></div></td>
            <td width="35%"> <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=Rs("login")%>&nbsp; 
              </font></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td width="20%" height="24" bgcolor="#EEEEEE"> <div align="right" background="../imagens/forum_images/table_bg2.gif"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Endere&ccedil;o:</strong></font></div></td>
            <td colspan="2"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%=Rs("endereco")%> &nbsp;<%=Rs("cidade")%> <%=Rs("estado")%> &nbsp;<%=Rs("cep")%></font></td>
            <td><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>E-mail:<a href="mailto:"> 
              </a></strong><a href="mailto:"><%=Rs("email")%></a></font></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td width="20%" bgcolor="#EEEEEE"> <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
                Nascimento .:</strong></font></div></td>
            <td width="16%"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%=Rs("niver")%> </font></td>
            <td width="16%"> <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Fone:</strong></font></div></td>
            <td width="35%"> <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=Rs("fone_res")%> 
              &nbsp;<strong>Cel.:</strong> <%=Rs("fone_cel")%> </font></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td width="20%" bgcolor="#EEEEEE"> <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
            <td width="16%"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font></td>
            <td width="16%"> <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Sexo:</strong></font></div></td>
            <td width="35%"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%=Rs("sexo")%> </font></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td width="13%" bgcolor="#EEEEEE">&nbsp;</td>
            <td width="20%" bgcolor="#EEEEEE"> <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Senha:</strong></font></div></td>
            <td width="16%"> <div align="left"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><font color="#000000"> 
                Encriptografado</font></font></div></td>
            <td width="16%">&nbsp;</td>
            <td width="35%">&nbsp;</td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td width="13%" bgcolor="#EEEEEE">&nbsp;</td>
            <td width="20%" bgcolor="#EEEEEE">&nbsp;</td>
            <td width="16%">&nbsp;</td>
            <td width="16%">&nbsp;</td>
            <td width="35%">&nbsp;</td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td bgcolor="#EEEEEE">&nbsp;</td>
            <td bgcolor="#EEEEEE"><strong><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></strong></td>
            <td colspan="3"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
              </font></td>
          </tr>
        </table>
  
  
  </td>
  </tr>
</table>
	 
  <!-- fIM tab -->
  <a href="#" onClick="window.close();"> ..:: Fechar ::.. </a> 
</center>
	 <%
	 Rs.Close
	 Set Rs=nothing
	 %>
	   