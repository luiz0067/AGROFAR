
<!-- #include file="login_nivel3.asp" -->


<%
Dim PageCurrente
Dim Rs
Dim I

if IsEmpty(Request.QueryString("PageNumber")) Then
  CurrentePage = 1
Else
  CurrentePage = cint(Request.QueryString("PageNumber"))
End If



Set Rs1 = Server.CreateObject("ADODB.Recordset")
Sqltext = "SELECT * FROM usuarios ORDER BY login ASC"
'where online = -1
Rs1.CursorLocation = 3
Rs1.Open SqlText, base_dados
Rs1.PageSize = 25
If Not Rs1.Eof Then
   Rs1.AbsolutePage = CurrentePage 
End IF
%>
<br>
<table width="70%" border="0" align="center" cellpadding="2" cellspacing="2">
  <tr> 
    <td><font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>P&aacute;gina 
      de Usu&aacute;rios</strong></font></td>
    <td><div align="right"><a href="usuarios_cadastra.asp"><img src="../imagens/outros_icos/botao_cadastrar_user.gif" alt="Cadastrar novo Usu&aacute;rio do Sistema" width="129" height="24" border="0"></a></div></td>
  </tr>
</table>

<br>

<table width="80%" border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center">
  <tr>
    <td>
	
	
      <table width="100%" border="0" cellpadding="1" cellspacing="0">
        <tr bgcolor="<%=cor_tabela%>"> 
          <td width="42" height="13" align="center" valign="top" bordercolor="#000000" background="../imagens/forum_images/table_bg.gif"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><span class="codigo"><b>ID</b></span></font></td>
          <td width="165" align="center" valign="top" background="../imagens/forum_images/table_bg.gif"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><span class="codigo"><b>Login/Nome</b></span></font></td>
          <td width="145" valign="top" background="../imagens/forum_images/table_bg.gif"> 
            <div align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><span class="codigo"><strong>Senha</strong></span></font></div></td>
          <td width="169" valign="top" background="../imagens/forum_images/table_bg.gif"><div align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><span class="codigo"><b>E-mail</b></span></font></div></td>
          <td width="188" valign="top" background="../imagens/forum_images/table_bg.gif"><div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>N&iacute;vel</strong></font></div></td>
          <td width="151" valign="top" background="../imagens/forum_images/table_bg.gif"> 
            <p align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>&nbsp;Ação</b></font></td>
        </tr>
        <% Do Until Rs1.AbsolutePage <> CurrentePage or Rs1.EOF

if flag = 0 then
	flag = 1
	bgcolor = "../imagens/outros_icos/bd_vb4.gif"
else
	flag = 0
	bgcolor = "../imagens/outros_icos/bd_vb5.gif"
end if
%>
        <tr valign="middle"> 
          <td width="42" height="13" align="center" bordercolor="#000000" background="<%= bgcolor %>" ><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%= rs1("id") %></font></td>
          <td width="165" align="center" background="<%= bgcolor %>"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%= rs1("login") %></font></td>
          <td width="145" background="<%= bgcolor %>"> <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Encriptada</font></div></td>
          <td width="169" background="<%= bgcolor %>"> <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs1("email") %></font></td>
          <td width="188" background="<%= bgcolor %>"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
		  
		  
		  <% If rs1("acesso") = "1" Then	%>
		  Básico
		  <% ElseIf rs1("acesso") = "2" Then	%>
		  Limitado
		  <%	ElseIf rs1("acesso") = "3" Then	%>
		  Administrador
		  <%	End If	%>
		  
		  
		  
		  </font></div></td>
          <td width="151" background="<%= bgcolor %>"><div align="center"><img src="../imagens/outros_icos/sm-editimage.gif" width="25" height="20" border="0" style="cursor:hand"  onclick="mostra_mais(<%=rs1("id")%>);" > 
              <a href="usuarios_editar.asp?id=<%= RS1("ID") %>" > <img src="../imagens/outros_icos/sm-edit.gif" alt="EDITAR" width="23" height="22" border="0"></a> 
              <a href="usuarios_deletar_admin.asp?id=<%= RS1("ID") %>" 
		  onClick="return confirm('TEM CERTEZA QUE DESEJA DELETAR??? \n ________________________________\n\n Esta ação irá deletar também todos \n os registros desse usuário \n\n OK-Para Deletar \n\n');"><img src="../imagens/outros_icos/sm-delete.gif" alt="DELETAR" width="20" height="20" border="0"></a> </font></div>
            </div></td>
        </tr>
        <%
Rs1.MoveNext
Loop

%>
      </table>
</td>
  </tr>
</table>
 <script>
	  function mostra_mais(id)	{
	  pg = "usuarios_ver_detalhe.asp?id=" + id
		var arr = showModalDialog(pg,window,"font-family:Verdana; font-size:12; status:false; dialogWidth:56; dialogHeight:20" );
	}
	</script>
	  
<!-- fIM tab -->

      <br><!--#include file="ativos_user.asp"-->
      
<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
          <td > 
            <%
		   if Rs1.PageCount > 1 Then ' Se a contagem de páginas tiver mais de 1 mostra os links     
	%>
            <div align="right"><font color="#999999" size="1" face="Verdana, Arial, Helvetica, sans-serif">Indicador 
              de Página [Mostrando <%=CurrentePage%> de <%=Rs1.PageCount%> páginas]</font></div>
            <table width="" height="1" border="1" align="right" bordercolor="#FFFFFF">
              <tr> 
                <td width="60" > 
                  <% 	If  not CurrentePage = 1 Then
				  ant = CurrentePage - 1	%>
                  <%	if acao = "Ativar" Then  
				  link = "?mostrar=off-line&PageNumber="
				  Else
				  link = "?mostrar=on-line&PageNumber="
                  End If %><a href="<%=link%><%=ant%>">
                  <div align="right"><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif">&laquo; 
                    anterior</font> 
                    <%	End If	%>
                  </div></td>
                <%	for I = 1 to Rs1.PageCount	
				If CurrentePage = I Then
		%>
                <td valign="middle" bordercolor="#666666" bgcolor="#999999"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=I%></font> 
                </td>
                <%  Else 	%>
                <td ><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%	if acao = "Ativar" Then  
				  link = "?mostrar=off-line&PageNumber="
				  Else
				  link = "?mostrar=on-line&PageNumber="
                  End If %>
                  <a href="<%=link%><%=I%>"><%=I%></a> </font></td>
                <%	End If	%>
                <%	Next	%>
                <td width="60" > 
                  <% 	If CurrentePage < Rs1.PageCount Then
				  prox = CurrentePage + 1	%><a href="<%=link%><%=prox%>">
                  <font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif">pr&oacute;xima 
                  &raquo; </font> 
                  <%	End If	%>
                </td>
              </tr>
            </table>
            <%
	I = I - 1
	End If
%>
            </td>
        </tr>
      </table>
     
<% 
Rs1.Close
Set Rs1 = Nothing
'Conn.close
'Set Conn = Nothing
%>

<!--#include file="rodape.asp"-->
