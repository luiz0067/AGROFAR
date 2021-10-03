<!-- #include file="login_nivel2.asp" -->
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
if request("tabela") = "Dicas" Then
tabela = "Dicas"
else
tabela = "Noticias"
End If


Set Rs1 = Server.CreateObject("ADODB.Recordset")
Sqltext = "SELECT * FROM "& tabela &"  where online = " & online & " ORDER BY  id DESC"
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
      <td width="50%" bgcolor="#FFFFFF"> <p align="center"><font color="<%=cor_links%>"><b><font face="Arial" size="3">Página 
          Administradora de <%=tabela%> &nbsp; <%=situacao%></font></b></font></td>
      <td width="18%"> 
      <table class="table">
<tr>
  <td><div align="center"><a class="linque" href="javascript:location.href='news_form.asp?tabela=<%=tabela%>';">Cadastrar Nova Notícia</a></div></td>
</tr>
</table>
  </td>
<td width="32%">
        <%
if online = "TRUE" Then
tp = "off-line"
tp_e = "INATIVOS"
Else
tp = ""
tp_e = "ATIVOS"
End If
%>
   <table class="table">
<tr>
  <td><div align="center"><a class="linque" href="javascript:location.href='news_default_admin.asp?mostrar=<%=tp%>';">Mostrar Notícias <%=tp_e%></a></div></td>
</tr>
</table></td>
    </tr>
  </table>

  <form name="form1" method="post" action="news_atualizar_admin.asp?op=1&ativar=<%=ativar%>&tabela=<%=tabela%>">
    <div align="center">
	<table width="790" border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center" >
                    <tr> 
                      <td>
      <table width="100%" border="0" cellpadding="2" cellspacing="0" >
        <tr bgcolor="<%=cor_tabela%>"> 
          <td width="33" height="13" align="center" valign="top" bordercolor="#000000" background="../imagens/icos_e_fundos/table_bg.gif"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><span class="codigo"><b>ID</b></span></font></td>
          <td width="66" align="center" valign="top" background="../imagens/icos_e_fundos/table_bg.gif"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><span class="codigo"><b>Data</b></span></font></td>
          <td width="391" valign="top" background="../imagens/icos_e_fundos/table_bg.gif"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><span class="codigo"><b>T&iacute;tulo</b></span></font></td>
          <td width="223" valign="top" background="../imagens/icos_e_fundos/table_bg.gif"> 
            <p align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>&nbsp;Açao</b></font></td>
          <td width="55" valign="top" background="../imagens/icos_e_fundos/table_bg.gif"> 
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
          <td width="33" height="13" align="center" background="<%= bgcolor %>" bordercolor="#000000"><font size="1" face="Arial, Helvetica, sans-serif"><a href="" class="codigo"><%= rs1("id") %></a></font></td>
          <td width="66" align="center" background="<%= bgcolor %>" ><font size="1" face="Arial, Helvetica, sans-serif"><span class="codigo"><%= rs1("data") %></span></font></td>
          <td width="391" background="<%= bgcolor %>" ><font size="1" face="Arial, Helvetica, sans-serif"><%= rs1("titulo") %></font></td>
          <td width="223" background="<%= bgcolor %>" > <div align="center"> <font face="Arial, Helvetica, sans-serif" size="1"> 
              <a href="news_deletar_admin.asp?id=<%= RS1("ID") %>&OP=1&mostrar=<%=request.QueryString("mostrar")%>&tabela=<%=tabela%>">[Deletar]</a> 
              <a href="news_form.asp?editar=ok&id=<%= RS1("ID") %>&tabela=<%=tabela%>">[Editar]</a> 
			  
			  <%
			  total = 0

			  Set Rs = Server.CreateObject("ADODB.Recordset")
			  Sqltext = "SELECT * FROM "&tabela&"_comentario WHERE id_noticia = '" & RS1("id") & "' order by id DESC"
				Rs.Open SqlText, base_dados
			  'conta = "SELECT * from noticias_comentario where id_noticia = " & RS1("ID")
			'Conn.Execute(conta)

If Rs.EOF AND Rs.BOF = TRUE Then
'nada				
Else
Do Until Rs.EOF
total = total + 1
Rs.MoveNext
Loop
Rs.Close
set Rs = nothing
End If

%>
			  <a href="news_comentarios_admin.asp?id=<%= RS1("ID") %>&tabela=<%=tabela%>">[comentários (<%=total%>) ]</a>
              </b></font> </div></td>
          <td width="55" background="<%= bgcolor %>" > <div align="center"> <font size="1" face="Arial, Helvetica, sans-serif"> 
              <input type="checkbox" name="<%= RS1("ID") %>" value="<%= RS1("ID") %>">
              </font></div></td>
        </tr>
     
	  <%
Rs1.MoveNext
Loop
%> </table>

</td></tr></table>
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
      <p>
        <input type="submit" name="Submit" value="Atualizar" style="background-color: <%=cor_fundo_box%>">
      </p>
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
