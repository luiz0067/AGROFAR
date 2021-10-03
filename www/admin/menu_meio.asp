
<!-- #include file="login_nivel2.asp" -->
<div align="center">
	<center>
	
  <table border="0" width="98%" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="23%" valign="top"> 
      
            <!-- Inicio da tabela de dados base -->
         
          <table   width="200"  border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center" >
            <tr> 
              <td> 
                <!-- Borda da Tabela -->
                <table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#FFFFFF"  >
                  <tr> 
                    <td background="../imagens/forum_images/table_bg2.gif"><div align="center"><font color="#336699" size="2" face="Arial, Helvetica, sans-serif"><strong>SUPORTE</strong></font></div></td>
                  </tr>
                  <tr> 
                    <td  bgcolor="#999999"></td>
                  </tr>
                  <tr> 
                    <td  bgcolor="#999999"></td>
                  </tr>
                  <tr> 
                    <td><p align="center"></td>
                  </tr>
                  <tr> 
                    <td  bgcolor="#999999"></td>
                  </tr>
                  <td> </td>
                  </tr>
                </table></td>
            </tr>
          </table>
          <!-- Fim de Borda e da Tabela Inicial -->
         
            <!-- Inicio da tabela de dados base -->
         
          <table  width="200"   border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center" >
            <tr> 
              <td> 
                <!-- Borda da Tabela -->
                <table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#FFFFFF"  >
                  <tr> 
                    <td background="../imagens/forum_images/table_bg2.gif"><div align="center"><font color="#336699" size="2" face="Arial, Helvetica, sans-serif"><strong>USU&Aacute;RIOS 
                        ONLINE </strong></font></div></td>
                  </tr>
                  <tr> 
                    <td  bgcolor="#999999"></td>
                  </tr>
                  <tr> 
                    <td  bgcolor="#999999"></td>
                  </tr>
                  <tr> 
                    <td><p align="center">
                        <!-- Mostra Users Online -->
                      </p>
                      <table width="95%" border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center">
                        <tr> 
                          <td> <table width="100%" cellpadding="3" cellspacing="0">
                              <tr bgcolor="#C0C0C0"> 
                                <td width="34%" background="../imagens/forum_images/table_bg.gif"> 
                                  <center>
                                    <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Usu&aacute;rio</b></font> 
                                  </center></td>
                                <td width="32%" background="../imagens/forum_images/table_bg.gif"> 
                                  <center>
                                    <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Online 
                                    a</b></font> </center></td>
                                <td width="34%" background="../imagens/forum_images/table_bg.gif"> 
                                  <center>
                                    <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Checado</b></font> 
                                  </center></td>
                              </tr>
                              <%
count=0
StrSql="Select * From active_users"
Set rs = Conn.execute (StrSql)
do while not rs.eof
count=count+1
if flag = 0 then
	flag = 1
	bgcolor = "../imagens/outros_icos/bd_vb4.gif"
else
	flag = 0
	bgcolor = "./imagens/outros_icos/bd_vb5.gif"
end if
%>
                              <tr bgcolor="#ffffff"> 
                                <td background="<%= bgcolor %>"><center>
                                    <font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs("nome")%></font> 
                                  </center></td>
                                <td background="<%= bgcolor %>"><center>
                                    <font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=datediff("n",rs("logon_time"),now())%> 
                                    Mins </font> </center></td>
                                <td background="<%= bgcolor %>"><center>
                                    <font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=datepart("h",rs("last_seen"))%>:<%=datepart("n",rs("last_seen"))%>:<%=datepart("s",rs("last_seen"))%> 
                                    </font> </center></td>
                              </tr>
                              <%
rs.movenext
loop
fechars
%>
                            </table></td>
                        </tr>
                      </table>
                      <div align="center"><font color="#FF0000" size="1" face="Arial, Helvetica, sans-serif"><strong>Atualmente 
                        est&#259;o <%=count%> usu&aacute;rio(s) on-line. <br>
						<% =application("nUsuarios") %> Pessoas navegando no site.</strong></font> 
                        <!--  Fim mostra user online -->
                      </div>
                      <p align="center">&nbsp; </p></td>
                  </tr>
                  <tr> 
                    <td  bgcolor="#999999"></td>
                  </tr>
                  <td> </td>
                  </tr>
                </table></td>
            </tr>
          </table>
          <!-- Fim de Borda e da Tabela Inicial -->
         </td>
      <td width="7%">&nbsp; &nbsp; </td>
      <td width="70%" valign="top"> 
        <p align="center"><font color="#003366" face="Arial, Helvetica, sans-serif" size="4">P&aacute;gina 
          Administradora do Site da <%=site_titulo%></font> 
        <center>
        <p align="justify"><font face="Arial, Helvetica, sans-serif" size="3" color="#000000"> 
          <img src="../<%=session("foto")%>" width="64" height="64" align="left" class="img_border">Ol&aacute; 
          <b><i><%=session("Nome")%></i></b>, voc&#281; est&aacute; logado com n&iacute;vel <b><%=session("Acesso_escreve")%></b></font> 
          <font face="Arial, Helvetica, sans-serif"><br>
          <font color="#000000">Nosso sistema iterativo esta cheio de novidades. 
          Muitas mesmo.<br>
          S&#259;o inumeras ferramentas de trabalho que tendem a melhor muito a rela&ccedil;&#259;o 
          com os s&oacute;cios trabalhando em um sistema de "pier to pier", ou seja, 
          um contato direto totalmente online entre todos os s&oacute;cios, de maneira 
          simples e f&aacute;cil, e o melhor, que presta um servi&ccedil;o ao s&oacute;cio, colocando-lhe 
          uma gama de informa&ccedil;&#337;es e deixando todos conectatos e trabalhando em 
          prol do <%=site_titulo%>. <br>
          <br>
          Para iniciar, solicitamos que atualize seus dados de cadastro, pois 
          nesta sess&#259;o existem inumeros recursos que s&#259;o se suma import&#259;ncia para 
          a funcionalidade do sistema e total intera&ccedil;&#259;o. <a href="socios_perfil.asp?socio=editar">Clique 
          aqui para atualizar/editar seus dados.</a><br>
          <br>
          Voc&ecirc; tamb&eacute;m pode acessar esta fun&ccedil;&atilde;o a qualquer 
          momento, ela fica no rodap&eacute; de todas as p&aacute;ginas, basta 
          clicar em &quot;Editar meu Perfil&quot;. Entre muitos recursos, voc&ecirc; 
          pode cadastrar seu skype, e assim que um s&oacute;cio se logar no site, 
          ele saber&aacute; que voc&ecirc; est&aacute; online, podendo teclar 
          com voc&ecirc; a qualquer momento.<br>
          <br>
          Caso precise de ajuda para usar o sistema, n&atilde;o existe em chamar 
          o suporte.<br>
          <br>
          Abra&ccedil;os a todos, espero que gostem do sistema.<br>
          <br>
          Luiz Fernando</font></font></p>
        </center>
		
         </td>
      <td width="0%">&nbsp;</td>
    </tr>
  </table>
	
<%
Conn.Close
	set Conn = nothing
	%>
<!--#include file="rodape.asp"-->
