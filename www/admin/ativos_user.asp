
<center><table width="80%" border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center">
  <tr>
    <td>
	
  <table width="100%" cellpadding="3" cellspacing="0">
          <tr bgcolor="#C0C0C0"> 
            <td background="imagens/forum_images/table_bg.gif"> <center>
                <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Usu&aacute;rio</b></font> 
              </center></td>
            <td background="imagens/forum_images/table_bg.gif"> <center>
                <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Endere&ccedil;o 
                IP</b></font> </center></td>
            <td background="imagens/forum_images/table_bg.gif"> <center>
                <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>&Uacute;ltima 
                Checagem de Data</b></font> </center></td>
            <td background="imagens/forum_images/table_bg.gif"> <center>
                <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Tempo 
                no Site</b></font> </center></td>
            <td background="imagens/forum_images/table_bg.gif"> <center>
                <font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><font color="#FFFFFF">&Uacute;ltima 
                P&aacute;gina</font><b></font> </center></td>
            <td background="imagens/forum_images/table_bg.gif">&nbsp;</td>
          </tr>
          <%
count=0
StrSql="Select * From active_users"
Set rs = Conn.execute (StrSql)
do while not rs.eof
count=count+1
if flag = 0 then
	flag = 1
	bgcolor = "imagens/outros_icos/bd_vb4.gif"
else
	flag = 0
	bgcolor = "imagens/outros_icos/bd_vb5.gif"
end if
%>
          <tr bgcolor="#ffffff"> 
            <td background="<%= bgcolor %>"><center>
                <font face="Verdana, Arial, Helvetica, sans-serif" size="1">User 
                <%=count%> - <%=rs("nome")%></font> </center></td>
            <td background="<%= bgcolor %>"><center>
                <font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs("User_Ip")%></font> 
              </center></td>
            <td background="<%= bgcolor %>"><center>
                <font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs("last_seen")%></font> 
              </center></td>
            <td background="<%= bgcolor %>"><center>
                <font face="Verdana, Arial, Helvetica, sans-serif" size="1"> <%=datediff("n",rs("logon_time"),now())%> 
                Mins </font> </center></td>
            <td background="<%= bgcolor %>"><center>
                <font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs("last_page")%></font> 
              </center></td>
            <td background="<%= bgcolor %>"><div align="center"><a href="activeusers_delete.asp?id=<%= RS("ID") %>" 
		  onClick="return confirm('TEM CERTEZA QUE DESEJA DELETAR??? \n ________________________________\n\n Esta ação irá deletar também todos \n os registros desse usuário \n\n OK-Para Deletar \n\n');"><img src="../imagens/outros_icos/xsm-delete.gif" alt="DELETAR Log desse Usuário" width="10" height="10" border="0"></a></div></td>
          </tr>
          <%
rs.movenext
loop
fechars
%>
        </table>
  </td>
  </tr>
</table>
  <font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Atualmente 
  estão <%=count%> usuário(s) on-line. </strong></font> 
</center>
