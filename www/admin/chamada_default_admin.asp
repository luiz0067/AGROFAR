<!-- #include file="login_nivel2.asp" -->

<a class="l">
<%
 Set Conn = Server.CreateObject("ADODB.Connection")
 Conn.Open base_dados

Dim PageCurrente
Dim Rs
Dim I

if IsEmpty(Request.QueryString("PageNumber")) Then
  CurrentePage = 1
Else
  CurrentePage = cint(Request.QueryString("PageNumber"))
End If



Set Rs1 = Server.CreateObject("ADODB.Recordset")
Sqltext = "SELECT * FROM chamada"
'where online = -1
Rs1.CursorLocation = 3
Rs1.Open SqlText, base_dados
Rs1.PageSize = 15
If Not Rs1.Eof Then
   Rs1.AbsolutePage = CurrentePage 
End IF

%>
<div align="center"> <b><font face="Arial" size="2" color="#000000">..:: Administrador 
  de Chamadas Topo de P&aacute;gina::.. <br>
  </font><font face="Arial" size="2"><br>
   <%	if request("msg") = "1" Then	%>
<table width="350" border="0" cellspacing="1" cellpadding="1" align="center" class="moldura1" height="30">
  <tr>
    <td><div align="center"><img src="../imagens/icos_e_fundos/warning.png" align="absmiddle"> <span style="color: #FF0000; font-weight: bold">
      Sea Chamada CADASTRADA com Sucesso!</span></div></td>
  </tr>
</table>
<br>
    <%	End If	%>
   <%	if request("msg") = "2" Then	%>
<table width="350" border="0" cellspacing="1" cellpadding="1" align="center" class="moldura1" height="30">
  <tr>
    <td><div align="center"><img src="../imagens/icos_e_fundos/warning.png" align="absmiddle"> <span style="color: #FF0000; font-weight: bold">
      Sua Chamada ATUALIZADA com Sucesso!</span></div></td>
  </tr>
</table>
<br>
    <%	End If	%>
     <%	if request("msg") = "3" Then	%>
<table width="350" border="0" cellspacing="1" cellpadding="1" align="center" class="moldura1" height="30">
  <tr>
    <td><div align="center"><img src="../imagens/icos_e_fundos/warning.png" align="absmiddle"> <span style="color: #FF0000; font-weight: bold">
      Sua Chamada DELETADA com Sucesso!</span></div></td>
  </tr>
</table>
<br>
    <%	End If	%>
    
<table class="table">
<tr>
  <td><div align="center"><a class="linque" href="javascript:location.href='chamada_cadastrar_admin.asp';">Adicionar Nova Chamada</a></div></td>
</tr>
</table>

  <form name="form1" method="post" action="frases_atualizar_admin.asp?op=1">
    <div align="center">
    <table border="1" bordercolor="cococo" cellspacing="0" cellpadding="0" width="684">
    <tr><td>
      <table width="684" border="0" cellpadding="0" cellspacing="0">
        <tr bgcolor="#000000"> 
          <td width="44" align="center" valign="top" background="../imagens/icos_e_fundos/table_bg_image.gif" bgcolor="#000000"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><span class="codigo"><b>ID</b></span></font></td>
          <td align="center" valign="top" background="../imagens/icos_e_fundos/table_bg_image.gif"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><span class="codigo"><b>Chamada</b></span></font></td>
          <td width="98" valign="top" background="../imagens/icos_e_fundos/table_bg_image.gif" bgcolor="#000000"> 
          <p align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>&nbsp;Ação</b></font></td>
        </tr>
        <% Do Until Rs1.AbsolutePage <> CurrentePage or Rs1.EOF

if flag = 0 then
	flag = 1
	bgcolor = "#EFEFEF"
else
	flag = 0
	bgcolor = "#FFFFCC"
end if

%>
        <tr valign="middle"class="l" > 
          <td width="44" align="center" bgcolor="<%= bgcolor %>" > 
            <font size="1" face="Arial, Helvetica, sans-serif"><a href="" class="l"> 
            <%= rs1("id") %></a> </font></td>
          <td align="center" bgcolor="<%= bgcolor %>"> <font size="1" face="Arial, Helvetica, sans-serif"><span class="codigo"><%= rs1("chamada") %></span></font> 
          </td>
          <td width="98" bgcolor="<%= bgcolor %>"> <div align="center"> <font color="#000000" size="1" face="Arial, Helvetica, sans-serif"> 
              <a href="chamada_deletar_admin.asp?id=<%= RS1("ID") %>&OP=1" class="l"
               OnClick="return confirm('Você tem certeza que você quer excluir esta Chamada?\n Esta ação não poderá ser desfeita!!!\n')"
              >[Deletar]</a> 
              <a href="chamada_editar_admin.asp?id=<%= RS1("ID") %>&frases=editar" class="l">[Editar]</a><b> 
              </b> </font> </div></td>
        </tr>
        <%
Rs1.MoveNext
Loop
%>
      </table>
      </td></tr></table>
		
      <p>&nbsp; 
      <table width="373" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="373" height="12" valign="top"> 
		  <div align="center"> <font color="#000000"> 
              <% For I = 1 to Rs1.PageCount%>
              <font color="#FF4646" size="2" face="Verdana, Arial, Helvetica, sans-serif">Páginas 
              </font><font color="#C40000"><a href="chamada_default_admin.asp?PageNumber=<%=I%>" class="l"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">[<%=I%>]</font></a></font> 
              <font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
              <% Next %>
              </font> </font> </div>
          </td>
          </tr>
      </table>
      <p>&nbsp; </p>
    </div>
  </form>
</div>
<% Rs1.Close
   Set Rs1 = Nothing
%>
<p align="center">
  <!--#include file="rodape.asp"-->
</p></a>
</body>
</html>



