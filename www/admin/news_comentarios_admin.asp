<!-- #include file="login_nivel2.asp" -->
<!-- #include file=".\config_site.inc" -->
<%
if request("tabela") = "Dicas" Then
tabela = "Dicas"
else
tabela = "Noticias"
End If
set Rs = server.createobject("ADODB.recordset")

if IsEmpty(Request.QueryString("PageNumber")) Then
  		CurrentePage = 1
  	Else
  		CurrentePage = cint(Request.QueryString("PageNumber"))
  	End If

Sqltext = "SELECT * FROM "&tabela&"_comentario WHERE id_noticia = '" & Request.QueryString("id") & "' order by id DESC"
Rs.CursorLocation = 3
Rs.Open SqlText, base_dados


if Rs.EOF AND Rs.BOF =True Then				
'Do Until Rs.AbsolutePage <> CurrentePage or Not Rs.EOF
%>
<b> <font color="#006600" size="1" face="Verdana, Arial, Helvetica, sans-serif">Nenhum 
comentário enviado para a <%=tabela%> </font><font color="#006600" size="2" face="Arial, Helvetica, sans-serif"><strong><%=Request.QueryString("id")%></strong></font><font color="#006600" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
</font></b></font> 
<%
Else
%>
<p></p>
<div align="center"><font color="#006600" size="2" face="Arial, Helvetica, sans-serif"><strong>Comentários 
  enviados para a <%=tabela%> <%=Request.QueryString("id")%></strong></font>
  <hr align="left" width="100%" noshade color="#CCCCCC">
  <Br>
  <%
Rs.PageSize = 20 ' Quatos números de registro será mostrado!!!
Rs.AbsolutePage = CurrentePage
Do Until Rs.AbsolutePage <> CurrentePage or  Rs.EOF				
%>
</div>
<p></p>
<table width="70%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="80%"><font color="#003399" size="1" face="Arial, Helvetica, sans-serif"><b><a href="mailto:<%=Rs("com_mail")%>"><%=Rs("com_nome")%> 
      <img src="../imagens/botoes/e-mail_topic.gif" alt="Clique aqui para enviar um e-mail para <%=Rs("com_mail")%>" width="16" height="11" border="0"></a></b></font> 
      - <font color="#cococo" size="1" face="Arial, Helvetica, sans-serif"><%=Rs("com_data")%></font> 
      <br>
      <font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><%=Rs("com_texto")%> 
      </font></td>
    <td align="center" valign="middle"> 
      <div align="center"><font color="#006600" size="2" face="Arial, Helvetica, sans-serif">[editar] 
        - [<a href="news_comentarios_deletar_admin.asp?id=<%=rs("id")%>&id_noticia=<%=Request.QueryString("id")%>">deletar</a>]</font></div></td>
  </tr>
  <tr> 
    <td colspan="2"><hr align="left" width="100%" noshade color="#CCCCCC"></td>
  </tr>
</table>

<% 
   Rs.MoveNext
   Loop
%>
<%  
	'For I = 1 to Rs.PageCount
	if Rs.PageCount > 1 Then
	
	I =  CurrentePage + 1
		if Rs.PageCount = CurrentePage Then
		' nao faz nada
		else
	%>
<font color="#003399" size="1" face="Arial, Helvetica, sans-serif"><acronym title="Clique aqui para ver mais comentários"><a href="noticias_ler.asp?pag=4&noticias=ler&id=<%=id%>&PageNumber=<%=I%>">ver 
mais comentários</a></acronym></font></b> 
<% 
		
				end if
				End If
            %>
<%
	'Rs.Close
    'set Rs = nothing
	'Conn.Close
	'set Conn = nothing
End If 
%>
<p></p>
<div align="center"><a href="news_default_admin.asp?tabela=<%=tabela%>"><font size="2" face="Arial, Helvetica, sans-serif">Voltar 
  para not&iacute;cias</font></a>
  <!--#include file="rodape.asp"-->
</div>
