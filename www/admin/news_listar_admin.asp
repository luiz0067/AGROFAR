<!--#include file="nivel.asp"-->
<% Response.Expires = -1 %>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.codigo {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 11px; color: #000000; text-decoration: none}
td {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; text-decoration: none}
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<div align="center">
<!--#include file="news_menu_admin.asp" -->
</div>
<!--#include file=".\config_site.inc"-->

<%
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open dominiundados
Sqltext = "SELECT * FROM messages ORDER BY ndate DESC, id DESC"
Set Rs = Conn.Execute(Sqltext)

%>
<div align="center"> 
  <p>Lista de Todas as Not&iacute;cias<br>
  </p>
  <form name="form1" method="post" action="news_deletar_admin.asp?OP=2&Pag=Listar.asp">
    <div align="center">
      <table width="600" border="0" cellpadding="0" cellspacing="1">
        <tr> 
          <td width="68" height="13" valign="top" bgcolor="#00CCFF" bordercolor="#000000" align="center"><span class="codigo">ID</span></td>
          <td width="111" valign="top" bgcolor="#00CCFF" align="center"><span class="codigo">Data</span></td>
          <td width="349" valign="top" bgcolor="#00CCFF"><span class="codigo">T&iacute;tulo</span></td>
          <td width="67" valign="top" bgcolor="#00CCFF"> 
            
          <div align="center" class="codigo">Deletar</div>
          </td>
        </tr>
      </table>

<% While NOT Rs.EOF
if flag = 0 then
	flag = 1
	bgcolor = "#66FFFF"
else
	flag = 0
	bgcolor = "white"
end if
%>
      
    <table width="600" border="0" cellpadding="0" cellspacing="1">
      <tr valign="middle"> 
        <td width="68" height="18" align="center" bgcolor="<%= bgcolor %>" bordercolor="#000000"><a href="" class="codigo"><%= rs("id") %></a></td>
        <td width="111" align="center" bgcolor="<%= bgcolor %>"><span class="codigo"><%= rs("NDate") %></span></td>
        <td width="353" valign="top" bgcolor="<%= bgcolor %>"><span class="codigo"><%= rs("Title") %></span></td>
        <td width="64" bgcolor="<%= bgcolor %>"> 
            <div align="center"> 
              <input type="checkbox" name="<%= RS("ID") %>" value="<%= RS("ID") %>">
            </div>
        </td>
      </tr>
    </table>
	  <%
Rs.MoveNext
Wend
%>
      <table width="600" border="0" cellpadding="0" cellspacing="1">
        <tr> 
          <td width="68" height="13" valign="top" bgcolor="#00CCFF" bordercolor="#000000" align="center"><span class="codigo">ID</span></td>
          <td width="111" valign="top" bgcolor="#00CCFF" align="center"><span class="codigo">Data</span></td>
          <td width="349" valign="top" bgcolor="#00CCFF"><span class="codigo">T&iacute;tulo</span></td>
          <td width="67" valign="top" bgcolor="#00CCFF"> 
            
          <div align="center" class="codigo">Deletar</div>
          </td>
        </tr>
      </table>
      
<p>
        <input type="submit" name="Submit" value="Deletar">
      </p>
    </div>
  </form>
</div>  
<%
   Rs.Close
   Set Rs = Nothing
%>
</body>
</html>
