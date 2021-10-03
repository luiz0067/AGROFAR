<!-- #include file="login_nivel2.asp" -->

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
Sqltext = "SELECT * FROM livro  ORDER BY ID DESC"
'where online = -1
Rs1.CursorLocation = 3
Rs1.Open SqlText, base_dados
Rs1.PageSize = 10
If Not Rs1.Eof Then
   Rs1.AbsolutePage = CurrentePage 
End IF

%>
<div align="center">
	
  <table border="0" width="100%" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#FFFFFF"> <p align="center"><b><font face="Arial" size="3" color="#000000">Página 
          Administradora do Livro <%=situacao%></font></b></td>
    </tr>
  </table>


    <div align="center">
      <table width="750" border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center" >
                    <tr> 
                      <td>
    
  <table width="750" border="0" cellpadding="0" cellspacing="0">
    <tr bgcolor="<%=cor_tabela%>"> 
      <td width="31" height="13" align="center" valign="top" bordercolor="#000000" background="../imagens/icos_e_fundos/table_bg.gif"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><span class="codigo"><b>ID</b></span></font></td>
      <td width="100" align="center" valign="top" background="../imagens/icos_e_fundos/table_bg.gif"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><span class="codigo"><b>Data</b></span></font></td>
      <td width="210" valign="top" background="../imagens/icos_e_fundos/table_bg.gif"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><span class="codigo"><b>T&iacute;tulo</b></span></font></td>
      <td width="305" valign="top" background="../imagens/icos_e_fundos/table_bg.gif"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><strong>Mensagem</strong></font></td>
      <td width="104" valign="top" background="../imagens/icos_e_fundos/table_bg.gif"> 
        <p align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>&nbsp;Ação</b></font></td>
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
      <td width="31" height="21" align="center" background="<%= bgcolor %>"  bordercolor="#000000"><font size="1" face="Arial, Helvetica, sans-serif"><%= rs1("id") %></font></td>
      <td width="100" align="center" background="<%= bgcolor %>" ><font size="1" face="Arial, Helvetica, sans-serif"><span class="codigo"><%= rs1("data") %></span></font></td>
      <td background="<%= bgcolor %>" ><font size="1" face="Arial, Helvetica, sans-serif"><span class="codigo"><%= rs1("titulo") %> 
        </span></font></td>
      <td background="<%= bgcolor %>" ><font size="1" face="Arial, Helvetica, sans-serif"><span class="codigo"><%= left(rs1("mensagem"),90) %>...</span></font></td>
      <td width="104" background="<%= bgcolor %>" > <div align="center"> <font face="Arial, Helvetica, sans-serif" size="1"><b> 
          <a href="livro_deletar_admin.asp?id=<%= RS1("ID") %>&OP=1&mostrar=<%=request.QueryString("mostrar")%>" class="l">[Deletar]</a> 
          <a href="livro_editar_admin.asp?id=<%= RS1("ID") %>&livro=form" class="l">[Editar]</a> 
          </b></font> </div></td>
    </tr>
    <%
Rs1.MoveNext
Loop
%>
  </table></td></tr></table>
		<p>
      
<br>
      <table width="373" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="373" height="12" valign="top"> 
		  <div align="center"> <font size="2" face="Arial, Helvetica, sans-serif"> 
          <font color="#006600"> 
          <%
		   if Rs1.PageCount > 1 Then ' Se a contagem de páginas tiver mais de 1 mostra os links     
	%>
          <font color="#000000" size="1">Mostrando página <%=CurrentePage%> de 
          <%=Rs1.PageCount%> páginas. 
          <%	for I = 1 to Rs1.PageCount	
	
	If CurrentePage = I Then
	%>
          <strong>[<%=I%>]</strong></font></font></font> <font color="#000000" size="1" face="Arial, Helvetica, sans-serif"> 
          <% 
	Else
			if acao = "Ativar" Then
			%>
          [<a href="?mostrar=off-line&PageNumber=<%=I%>" class="l"><%=I%></a>] 
          <% else %>
          [<a href="?mostrar=on-line&PageNumber=<%=I%>" class="l"><%=I%></a>] 
          <%
			end if
	End If
 
	Next
	I = I - 1
	End If
              
%>
          </font><font size="1" face="Arial, Helvetica, sans-serif"></font> </font> 
        </div>
          </td>
          </tr>
      </table>
      
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
