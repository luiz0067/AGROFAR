<!-- #include file="login_nivel3.asp" -->
<!--#include file=".\config_site.inc" -->
<%
id = request("id")

set Rs = server.createobject("ADODB.recordset")

Sqltext = "SELECT * FROM noticias WHERE id = "& id 
Set Rs = Conn.Execute(Sqltext)
%>

<center>
  <font face="Arial" size="3" color="#000000"><b>:: Editando noticias e Artigos ::</b></font> 
</center> 
<center>
<% IF Request.Form("mensagem") = "" THEN %>
<table border="2" bordercolor="cococo" cellspacing="0" cellpadding="0" width="500"><tr><td width="403">

<FORM METHOD="POST" ACTION="?noticias=atualizar&id=<%=id%>" >
  <div align="center">
	    <table width="557" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#C0C0C0">
          <tr> 
            <td align="right" bgcolor="#CCCCCC" colspan="3"> <p align="center"><b><font face="Arial" size="2" color="#000000">EDITOR 
                DE INTERATIVO DE TEXTO</font></b> </td>
          </tr>
          <tr> 
            <td width="28%" align="right" bgcolor="#CCCCCC"><font size="2" face="Arial, Helvetica, sans-serif"><b>T&iacute;tulo:&nbsp;&nbsp;</b></font> 
            </td>
            <td width="72%" colspan="2" bgcolor="#CCCCCC"><font size="2" face="Arial, Helvetica, sans-serif"> 
              
            <input name="titulo" type="text" id="titulo" style="background-color: <%=cor_fundo_box%>" value="<%=Rs("titulo")%>" size="32">
              </font></td>
          </tr>
          <tr> 
            
          <td align="right" bgcolor="#CCCCCC"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Cabe&ccedil;a:</strong> 
            </font></td>
            
          <td colspan="2" bgcolor="#CCCCCC"><font size="2" face="Arial, Helvetica, sans-serif"> 
            <textarea name="cabeca" cols="35" rows="2" id="cabeca" style="background-color: <%=cor_fundo_box%>"><%=Rs("cabeca")%></textarea>
            </font></td>
          </tr>
          <tr> 
            <td width="28%" align="right" bgcolor="#CCCCCC"><font size="2" face="Arial, Helvetica, sans-serif"><b>Autor:&nbsp;&nbsp;</b></font> 
            </td>
            <td width="72%" colspan="2" bgcolor="#CCCCCC"><font size="2" face="Arial, Helvetica, sans-serif"> 
              
            <input name="autor" type="text" id="autor" style="background-color: <%=cor_fundo_box%>" value="<%=Rs("autor")%>" size="32">
              </font></td ????A????>
          </tr>
          <tr> 
            <td align="right" bgcolor="#CCCCCC"><font size="2" face="Arial, Helvetica, sans-serif"><b>E-mail 
              :&nbsp;</b></font></td>
            <td colspan="2" bgcolor="#CCCCCC"><font size="2" face="Arial, Helvetica, sans-serif"> 
              
            <input name="mailautor" type="text" id="mailautor" style="background-color: <%=cor_fundo_box%>" value="<%=Rs("mailautor")%>" size="32">
              </font></td>
          </tr>
          <tr> 
            <td width="28%" align="right" bgcolor="#CCCCCC"><font size="2" face="Arial, Helvetica, sans-serif"><b>Data 
              do artigo:&nbsp;</b></font> </td>
            <td width="72%" colspan="2" bgcolor="#CCCCCC"><font size="2" face="Arial, Helvetica, sans-serif"> 
              
            <input name="data" type="text" id="data" style="background-color: <%=cor_fundo_box%>" value="<%=Rs("data")%>" size="32">
              </font></td>
          </tr>
          <tr> 
            <td width="28%" align="right" bgcolor="#CCCCCC"> <font size="2" face="Arial, Helvetica, sans-serif"> 
              <b> Recomendado:&nbsp;&nbsp;</b></font> </td>
            <td width="72%" colspan="2" bgcolor="#CCCCCC"><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="recomendado" maxlength="2" size="2" value="<%=Rs("recomendado")%>" style="background-color: <%=cor_fundo_box%>" >
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
              <b>N&ordm; inicial de Cliques</b>: 
              <input type="text" name="contagem"  value="<%=Rs("contagem")%>" size="2" style="background-color: <%=cor_fundo_box%>">
              </font></td>
          </tr>
          <tr> 
            <td colspan="3" align="right" bgcolor="#CCCCCC"><textarea name="texto" cols="100" rows="30" id="texto"><%=Rs("texto")%></textarea></td>
          </tr>
          <tr> 
            <td colspan="3" align="right" bgcolor="#CCCCCC"><input type="submit" name="Submit" value=":: ATUALIZAR ::"></td>
          </tr>
          <tr> 
            <td></td>
          </tr>
        </table>
		</table>
	  
           </form>
	
  
<% 
	Rs.Close
	Set Rs = Nothing
	Conn.Close
   	Set Conn = Nothing
   'Rs.Close
   'Set Rs = Nothing

End If 

If Request.QueryString("noticias") = "atualizar" Then
   
   Set Conn = Server.CreateObject("ADODB.Connection")
   		Conn.Open base_dados  
   	      		   
	       Sql="UPDATE  noticias SET titulo = '" & Request.form("titulo") &  "',  autor = '" & Request.form("autor") &  "' , mailautor = '" & Request.form("mailautor") & "' , texto= '" & Request.form("texto") & "', data = '" & request.form("data") & " ',contagem = " & Request.form("contagem") &  ", recomendado = " & Request.form("recomendado") & ", cabeca = '" & Request.form("cabeca") & "' WHERE id =" & request.QueryString("id")     		
	       Conn.Execute(Sql)
	      		 
	      'aee

   %>
    
<script>
		msg = "           DADOS ATUALIZADOS COM SUCESSO!!!         \n";
		msg += "_____________________________________\n\n";
		
		msg += "Os dados da notícia '<%=Request.form("titulo")%>' foram atualizados com sucesso!\n\n";
		msg += "A Equipe do <%=site_titulo%> agradece a sua contribuição. \n";
		
		alert(msg + "\n\n");
		location.href='noticias_default_admin.asp';
		
		
		</script>

   <%	
   
   Conn.Close
   	Set Conn = Nothing
   end if
  
  
%>

