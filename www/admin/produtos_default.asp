<!-- #include file="login_nivel3.asp" -->
<%

If request.QueryString("mostrar") = "off-line" Then
online = "FALSE"
ativar = "sim"
situacao = "OFF-LINE"
acao  = "Ativar"
else
online = "TRUE"
situacao = "ON-LINE"
acao = "Desativar"
End If

%>
<% response.expires = -1 %>



<!--#include file=".\config_site.inc"-->
<title>Administrador do site da <%=site_titulo%> - Desenvolvido por <%=desenvolvedor%></title>
<%
 'Set Conn = Server.CreateObject("ADODB.Connection")
 'Conn.Open gruber

Dim PageCurrente
Dim Rs
Dim I

if IsEmpty(Request.QueryString("PageNumber")) Then
  CurrentePage = 1
Else
  CurrentePage = cint(Request.QueryString("PageNumber"))
End If

Set Rs1 = Server.CreateObject("ADODB.Recordset")
Sqltext = "SELECT * FROM produtos  where online = " & online & " ORDER BY  id DESC"
'where online = -1
Rs1.CursorLocation = 3
Rs1.Open SqlText, base_dados
Rs1.PageSize = 20
If Not Rs1.Eof Then
   Rs1.AbsolutePage = CurrentePage 
End IF

%>
<div align="center">
	
  <table border="0" width="100%" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="55%" bgcolor="#FFFFFF"> <p align="center"><font color="<%=cor_links%>"><b><font face="Arial" size="3">Página 
          Administradora de Produtos &nbsp; <%=situacao%></font></b></font></td>
      <td width="45%"> <font color="<%=cor_links%>" size="2" face="Arial, Helvetica, sans-serif"><b> 
        <input name="button"  type="button" style="cursor:hand;font-size:9pt;color: #CC0000;border-width:2px ;width:180px;"   onClick="location.href='produto_cadastrar.asp?ac=cadastrar';"  value="Cadastrar Novos Produtos ">
        <%
if online = "TRUE" Then
%>
        <input name="button2"  type="button" style="cursor:hand;font-size:9pt;color: #CC0000;border-width:2px ;width:180px;"  onClick="location.href='?mostrar=off-line';"  value="Mostrar Produtos  INATIVOS" title="Clique Aqui para Mostrar os Produtos OFF-LINE, Produtos Inativos">
        <% Else  %>
        <input name="button2"  type="button" style="cursor:hand;font-size:9pt;color: #CC0000;border-width:2px ;width:180px;"   onClick="location.href='produtos_default.asp';"  value="Mostrar Produtos  ATIVOS">
        <% End If %>
        </b></font></td>
    </tr>
  </table>

  <form name="form1" method="post" action="produtos_atualizar.asp?op=1&ativar=<%=ativar%>">
    <div align="center">
	<table width="790" border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center" >
                    <tr> 
                      <td>
      <table width="100%" border="0" cellpadding="2" cellspacing="0" >
              <tr bgcolor="<%=cor_tabela%>"> 
                <td width="36" height="13" align="center" valign="top" bordercolor="#000000" background="../imagens/forum_images/table_bg.gif"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><span class="codigo"><b>ID</b></span></font></td>
                <td width="161" valign="top" background="../imagens/forum_images/table_bg.gif"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><span class="codigo"><b>T&iacute;tulo 
                  no Menu</b></span></font></td>
                <td width="171" valign="top" background="../imagens/forum_images/table_bg.gif"><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">T&iacute;tulo 
                  na P&aacute;gina</font></strong></td>
                <td width="172" valign="top" background="../imagens/forum_images/table_bg.gif"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Sub 
                  T&iacute;tulo </strong></font></td>
                <td width="140" valign="top" background="../imagens/forum_images/table_bg.gif"> 
                  <p align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>&nbsp;Ação</b></font></td>
                <td width="84" valign="top" background="../imagens/forum_images/table_bg.gif"> 
                  <div align="center" class="codigo"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b><%=acao%></b></font></div></td>
              </tr>
              <% Do Until Rs1.AbsolutePage <> CurrentePage or Rs1.EOF

if flag = 0 then
	flag = 1
	bgcolor = "../imagens/outros_icos/bd_vb2.gif"
else
	flag = 0
	bgcolor = "../imagens/outros_icos/bd_vb3.gif"
end if
%>
              <tr valign="middle"> 
                <td width="36" height="13" align="center" background="<%= bgcolor %>" bordercolor="#000000"><font size="1" face="Arial, Helvetica, sans-serif"><%= rs1("id") %></font></td>
                <td width="161" background="<%= bgcolor %>" ><font size="1" face="Arial, Helvetica, sans-serif"><span class="codigo" 
				>
				<a href="javascript:showModalDialog('../pg_produtos.asp?id=<%=rs1("id")%>',window,'font-family:Verdana; font-size:12; status:No; scrollbars:yes; dialogWidth:52; dialogHeight:30')"><%= rs1("tit_menu") %></a></span></font></td>
                <td width="171" background="<%= bgcolor %>" ><font size="1" face="Arial, Helvetica, sans-serif"><span class="codigo"><%= rs1("titulo") %></span></font></td>
                <td width="172" background="<%= bgcolor %>" ><font size="1" face="Arial, Helvetica, sans-serif"><span class="codigo"><%= rs1("subtitulo") %></span></font></td>
                <td width="140" background="<%= bgcolor %>" > <div align="center"><a href="news_deletar_admin.asp?id=<%= RS1("ID") %>&OP=1&mostrar=<%=request.QueryString("mostrar")%>&tabela=<%=tabela%>"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">[Deletar]</font></a> 
                    <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="produto_cadastrar.asp?ac=editar&id=<%= RS1("ID") %>">[Editar]</a> 
                    </font> </div>
                <td width="84" background="<%= bgcolor %>" > <div align="center"> 
                    <font size="1" face="Arial, Helvetica, sans-serif"> 
                    <input type="checkbox" name="<%= RS1("ID") %>" value="<%= RS1("ID") %>">
                    </font></div></td>
              </tr>
              <%
Rs1.MoveNext
Loop
%>
            </table>

</td></tr></table>
		
      
<br>
      <table width="373" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="373" height="12" valign="top"> 
		  <div align="center"> <font size="2" face="Arial, Helvetica, sans-serif"> 
              <font color="#006600"> 
              <%
		   if Rs1.PageCount > 1 Then ' Se a contagem de páginas tiver mais de 1 mostra os links     
	%>
              <font color="<%=cor_links%>"> Mostrando página <%=CurrentePage%> 
              de <%=Rs1.PageCount%> páginas. 
              <%	for I = 1 to Rs1.PageCount	
	
	If CurrentePage = I Then
	%>
              <strong>[<%=I%>]</strong></font></font></font> <font color="<%=cor_links%>" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <% 
	Else
			if acao = "Ativar" Then
			%>
              [<a href="?tabela=<%=tabela%>&mostrar=off-line&PageNumber=<%=I%>"><%=I%></a>] 
              <% else %>
              [<a href="?tabela=<%=tabela%>&mostrar=on-line&PageNumber=<%=I%>"><%=I%></a>] 
              <%
			end if
	End If
 
	Next
	I = I - 1
	End If
              
%>
              </font><font color="<%=cor_links%>"></font></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              </font> </div>
          </td>
          </tr>
      </table>

        <input type="submit" name="Submit" value="Atualizar" style="background-color: <%=cor_fundo_box%>">
   
    </div>
  </form>
</div>
<% 
Rs1.Close
Set Rs1 = Nothing
Conn.close
Set Conn = Nothing
%>
<p align="center">
  <!--#include file="rodape.asp"-->
</p>
</body>
</html>
