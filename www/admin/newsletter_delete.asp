<!-- #include file="login_nivel3.asp" -->
<!--#include file=".\config_site.inc" -->
<h2 align="center">
<hr color="#000000">
  <font size="3" face="Arial, Helvetica, sans-serif">Vendo e Deletando mebros 
  da Lista de E-mails<br></font>
  <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="newsletter_admin_menu.asp">Voltar 
  para página administradora de Newslettter</a></font> </h2></div>
<%

newsletter = "SELECT count(id) as total_newsletter from mail_list WHERE online=true" 
Set RS = Conn.Execute(newsletter)
total_newsletter = Rs("total_newsletter")
Rs.Close
Set Rs=Nothing

if IsEmpty(Request.QueryString("PageNumber")) Then
  CurrentePage = 1
Else
  CurrentePage = cint(Request.QueryString("PageNumber"))
End If

If request("mostrar") = "offline" Then
ordem = "False"
Else
ordem = "True"
End If

	Set Rs = Server.CreateObject("ADODB.Recordset")
		Sqltext = "SELECT * FROM mail_list WHERE online = " & ordem & " order by id DESC"
		
		Rs.CursorLocation = 3
		Rs.Open SqlText, base_dados
		Rs.PageSize = 30
		If Not Rs.Eof Then
   		Rs.AbsolutePage = CurrentePage 
		End IF
		If Rs.EOF AND Rs.BOF = TRUE Then
		%>
		
<P> <strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Nenhum 
  e-mail cadastrado no nosso banco de dados
  <% If request("mostrar") = "offline" Then	%>
      OffLine. <p></p><a href="newsletter_delete.asp">(clique aqui para ver os e-mails 
      ONLINE)</a> 
      <%	Else	%>
      OnLine. <p></p><a href="newsletter_delete.asp?mostrar=offline">(clique aqui para 
      ver os e-mails OFFLINE)</a> 
      <% 	End If	%></font></strong></P>				
		<%
		Else
		%>
		<form action="newsletter_delete_send.asp?mostrar=<%=Request("mostrar")%>&pg=<%=Currentepage%>" method="post">
<table width="50%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td colspan="4"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Mostrando 
      <%=Rs.PageSize%> e-mails de um total de <%=total_newsletter%> cadastrados:</strong></font> 
      <hr color="#000000"></td>
  </tr>
  <tr bgcolor="#FFCC00"> 
    <td colspan="4"><font color="#006600" size="1" face="Verdana, Arial, Helvetica, sans-serif">E-mails 
      <% If request("mostrar") = "offline" Then	%>
      OffLine <a href="newsletter_delete.asp">(clique aqui para ver os e-mails 
      ONLINE)</a> 
      <%	Else	%>
      OnLine <a href="newsletter_delete.asp?mostrar=offline">(clique aqui para 
      ver os e-mails OFFLINE)</a> 
      <% 	End If	%>
      </font></td>
  </tr>
  <%
		Do Until Rs.AbsolutePage <> CurrentePage or Rs.EOF

if flag = 0 then
	flag = 1
	bgcolor = "#cococo"
else
	flag = 0
	bgcolor = "white"
end if
					%>
  <tr bgcolor="<%=bgcolor%>"> 
    <td width="6%"> <div align="center"> 
          <input name="id" type="checkbox" value="<%=Rs("id")%>">
      </div></td>
    <td width="45%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=Rs("email")%></font></td>
    <%
					Rs.MoveNext
					Loop
					%>
  <tr> 
    <td colspan="4"><hr color="#000000"> </td>
  </tr>
</table>
  <div align="center">
    <input name="acao"  type="submit" id="acao" value="Deletar E-mails Selecionados">
    
	<% 
	If request("mostrar") = "offline" Then	
	escreve = "Online"
	else
	escreve = "Offline"
	End If
	%>
		<input name="acao"  type="submit" id="acao" value="Setar E-mails <%=escreve%>">
  </div>
</form>
					<%
					End If
					%>
					<table width="373" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          
  <td width="373" height="12" valign="top"> <div align="center"> <font size="2" face="Arial, Helvetica, sans-serif"> 
        <font color="#006600"> 
        <%
		   if Rs.PageCount > 1 Then ' Se a contagem de páginas tiver mais de 1 mostra os links     
	%>
        <font size="1" face="Verdana, Arial, Helvetica, sans-serif">Mostrando 
        página <%=CurrentePage%> de <%=Rs.PageCount%> páginas. 
        <%	for I = 1 to Rs.PageCount	
	
	If CurrentePage = I Then
	%>
        <strong>[<%=I%>]</strong></font></font></font> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
        <% 
	Else
			if Request("mostrar") = "offline" Then
			%>
        </font>[<a href="?mostrar=offline&PageNumber=<%=I%>"><%=I%></a>] 
        <% else %>
        [<a href="?PageNumber=<%=I%>"><%=I%></a>] 
        <%
			end if
	End If
 
	Next
	I = I - 1
	End If
              
%>
        </font></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        </font> </div>
          </td>
          </tr>
      </table>
	  <%
	  ' fecha tudo
	  Rs.Close
	  Set Rs=Nothing
	  %>