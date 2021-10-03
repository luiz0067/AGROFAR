<!-- #include file="login_nivel3.asp" -->


<!--#include file=".\config_site.inc"-->

<%
If Request.QueryString("livro") = "form" Then
id = request("id")

set Rs = server.createobject("ADODB.recordset")

Sqltext = "SELECT * FROM livro WHERE id = "& id 
Set Rs = Conn.Execute(Sqltext)
%>






<div align="center"></div>
<FORM METHOD="POST" ACTION="livro_editar_admin.asp?livro=atualizar&id=<%=id%>" > 
<table width="300" border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center" >
                    <tr> 
                      <td>
  
        
        
  <table  width="300"border="0" align="center" cellpadding="0" cellspacing="0" >
          <tr bgcolor="#F3F3F3"> 
            <td colspan="3" align="right" background="../imagens/forum_images/table_bg.gif"> 
              <div align="center"><font color="#FFFFFF"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Editando 
                Livro de Visitas </font></strong></font> </div></td>
          </tr>
          <tr bgcolor="#F3F3F3"> 
            <td width="100%" align="right"><strong><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><span class="titulo" id="var1">Titulo:&nbsp;</span></font></strong></td>
            <td colspan="2" nowrap> <strong><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"> 
              <input name="titulo" type="text" style="background-color: #c0c0c0; font-family: Arial; font-size: 8 pt; " value="<%=Rs("titulo")%>" size="35">
              </font> </strong></td>
          </tr>
          <tr bgcolor="#F3F3F3"> 
            <td width="100%" align="right"><strong><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><span class="titulo" id="var1">Autor:&nbsp;</span></font></strong></td>
            <td colspan="2" nowrap> <strong><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"> 
              <input name="nome" type="text" value="<%=Rs("nome")%>" id="nome" style="background-color: #c0c0c0; font-family: Arial; font-size: 8 pt; " size="35">
              </font></strong></td>
          </tr>
          <tr bgcolor="#F3F3F3"> 
            <td width="100%" align="right"> <strong><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><span class="titulo" id="var1">E-mail: 
              </span></font></strong></td>
            <td colspan="2"> <strong><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"> 
              <input name="email" type="text" value="<%=Rs("email")%>" id="email" style="background-color: #c0c0c0; font-family: Arial; font-size: 8 pt; " size="35">
              </font></strong></td>
          </tr>
          <tr bgcolor="#F3F3F3"> 
            <td width="100%" align="right"><strong><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Cidade: 
              </font></strong></td>
            <td colspan="2" valign="top" nowrap> <strong><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"> 
              <input name="cidade" type="text" value="<%=Rs("cidade")%>" id="cidade3" style="background-color: #c0c0c0; font-family: Arial; font-size: 8 pt; " size="21">
              UF: 
              <input name="estado" type="text"  value="<%=Rs("estado")%>" style="background-color: #c0c0c0; font-family: Arial; font-size: 8 pt; " size="2" maxlength="2">
              &nbsp;</font></strong></td>
          </tr>
          <tr bgcolor="#F3F3F3"> 
            <td align="right"><strong><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Mensagem:</font></strong></td>
            <td colspan="2" valign="top" nowrap> <strong><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"> 
              <textarea name="mensagem" cols="37" rows="6" id="mensagem" style="background-color: #c0c0c0; font-family: Arial; font-size: 8 pt; color: #006600" ><%=Rs("mensagem")%></textarea>
              &nbsp;</font></strong></td>
          </tr>
          <tr bgcolor="#F3F3F3"> 
            <td width="100%" align="right">&nbsp;</td>
            <td colspan="2" valign="top"><font color="#006600" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <input type="submit" name="Submit" value="   Atualizar   " style=" color: #ff0000">
              </font></td>
          </tr>
        </table>    
 </td></tr></table>
      </form>
<% 
	Rs.Close
	Set Rs = Nothing
	Conn.Close
   	Set Conn = Nothing
   
End If 

If Request.QueryString("livro") = "atualizar" Then


if Request.Form("publica_mail") = "sim" Then
	publica_mail = "true"
	Else
	publica_mail = "false"
End If

If Request.Form("newsletter") = "sim" Then
	newsletter = "true"
	else
	newsletter = "false"
End If   
   Set Conn = Server.CreateObject("ADODB.Connection")
   		Conn.Open base_dados  
   	      		   
	       Sql="UPDATE  livro SET titulo = '" & Request.form("titulo") &  "'," 'WHERE id =" & request.QueryString("id")
		   SQL = SQL + "nome = '" & Request.form("nome") &  "',"
		   SQL = SQL + "email = '" & Request.form("email") & "',"       		
	       SQL = SQL + "cidade= '" & Request.form("cidade") & "',"
		   SQL = SQL + "estado= '" & Request.form("estado") & "',"
		   SQL = SQL + "mensagem= '" & Request.form("mensagem") & "',"
		   SQL = SQL + "publica_mail= " & publica_mail & ",newsletter= " & newsletter & " "
		   SQL = SQL + "WHERE id =" & request.QueryString("id")
		   
		   Conn.Execute(Sql)
	      		 
	      'aee

   %>
    
<script>
		msg = "           DADOS ATUALIZADOS COM SUCESSO!!!         \n";
		msg += "_____________________________________\n\n";
		
		msg += "Os dados da notícia '<%=Request.form("titulo")%>' foram atualizados com sucesso!\n\n";
		msg += "A Equipe do <%=site_titulo%> agradece a sua contribuição. \n";
		
		alert(msg + "\n\n");
		location.href='livro_default_admin.asp';
		
		
		</script>
<%	
   
   Conn.Close
   	Set Conn = Nothing
   end if
  
  
%>
<p align="center">
  <!--#include file="rodape.asp"-->
