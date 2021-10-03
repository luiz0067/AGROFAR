<!--#include file="login_nivel2.asp"-->
<SCRIPT src="../css/tooltip.js" 
type=text/javascript></SCRIPT>
<%
if request("ac") = "Cadastrar" Then
ac="Cadastrar"
foto = request("temp_img")
else
ac="Editar"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "SELECT * FROM fotos where id="&request("id"), conn
	id_cat = rs("id_cat")
	titulo= rs("titulo")
	descricao = rs("descricao")
	autor = rs("autor")
	visto = rs("visto")
	foto = rs("foto")
	

End If
If request("foto") = "Cadastrar" Then
Set Conn = Server.CreateObject("ADODB.Connection")
   Conn.Open base_dados
 	Sql="INSERT INTO fotos (titulo,autor,mail,descricao,id_cat,foto,data,visto,online)"
	 sql = sql + "VALUES('" & Request("titulo") & "','" & Request("autor") & "','" & Request("mail") & "',"
	 sql = sql + "'" & Request("descricao") & "'," & Request("id_cat") & ",'" & Request("temp_img") & "','" & date & "',0,true)"
   Conn.Execute(Sql)
   %>
   <script>
   location.href='fotos_admin.asp?id_cat=<%=Request("id_cat")%>&amp;ver=categoria';
   </script>
   <%
   End If
   
   If request("foto") = "Editar" Then
   Set Conn = Server.CreateObject("ADODB.Connection")
   Conn.Open base_dados
   Sql = "UPDATE fotos SET titulo='" & Request("titulo") & "',autor='" & Request("autor") & "',"
   Sql = Sql + " descricao='" & Request("descricao") & "',data='" & date & "',"
   Sql = Sql + " id_cat = " & Request("id_cat") & ", foto='" & Request("temp_img") & "' WHERE Id =" & request("id")
  Conn.Execute(Sql)
   %>
   <script>
   location.href='fotos_admin.asp?id_cat=<%=Request("id_cat")%>&amp;ver=categoria';
   </script>
   <%
End If
%>
<script language="JavaScript">

  function verify(form){
        if(form.id_cat.value == ""){
		msg = "           CAMPO OBRIGATÓRIO         \n";
		msg += "_____________________________________\n";
		msg += "Selecione uma Categoria para cadastrar.\n";
		msg += "Este é uma campo Obrigatório\n";
		msg += "Você está sendo direcionado para o campo. \n";
		alert(msg);
            form.id_cat.focus();
            return false;
      }
	  if(form.titulo.value == ""){
		msg = "           CAMPO OBRIGATÓRIO         \n";
		msg += "_____________________________________\n";
		msg += "Preencha o campo [Título] para cadastrar.\n";
		msg += "Este é uma campo Obrigatório\n";
		msg += "Você está sendo direcionado para o campo. \n";
		alert(msg);
            form.titulo.focus();
            return false;
      }
	  if(form.autor.value == ""){
		msg = "           CAMPO OBRIGATÓRIO         \n";
		msg += "_____________________________________\n";
		msg += "Preencha o campo [Autor da Foto] para cadastrar.\n";
		msg += "Este é uma campo Obrigatório\n";
		msg += "Você está sendo direcionado para o campo. \n";
		alert(msg);
            form.autor.focus();
            return false;
      }if(form.descricao.value == ""){
		msg = "           CAMPO OBRIGATÓRIO         \n";
		msg += "_____________________________________\n";
		msg += "Preencha o campo [Descrição] para cadastrar.\n";
		msg += "Este é uma campo Obrigatório\n";
		msg += "Você está sendo direcionado para o campo. \n";
		alert(msg);
            form.descricao.focus();
            return false;
      }
}
</script>
	  <form  method="post" action="fotos_editar_admin.asp?foto=<%=request("ac")%>&id=<%=request("id")%>&ac=<%=request("ac")%>" onSubmit="return verify(this)">
	  <table width="500" border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center" height="170">
                          <tr> 
                            <td> 
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="2" bgcolor="#F3F3F3">
          <tr background="../imagens/forum_images/table_bg.gif" > 
            <td colspan="3"> 
              <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>:: 
            Cadastrar Foto ::</strong></font></div></td>
          </tr>
          <tr> 
            <td rowspan="5"><div align="center"><input name="temp_img" type="hidden" value="<%=request("temp_img")%>">
                <img src="../upload/fotos/p_<%=request("temp_img")%>.jpg" onmouseover="showtrail(150,150,'../upload/fotos/m_<%=request("temp_img")%>.jpg');" onmouseout=hidetrail();><br>
                <font size="1" face="Arial, Helvetica, sans-serif">Passe o mouse 
                para<br>
                ampliar a Foto</font></div></td>
            <td colspan="2"> <select  name="id_cat" style="font-family: Arial; font-size: 8 pt; border-style: solid; border-color: #000000;  WIDTH: 250px;">
                <option   value="" >&nbsp;..:: Selecione uma Categoria para Cadastrar::..</option>
				
                <%
		  Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "SELECT * FROM fotos_categorias order by id", conn

	Do Until Rs.EOF
	
			sqlcount = "SELECT count(titulo) FROM fotos where id_cat=" & rs("id") 
			Set RScounts = Conn.Execute(sqlcount)
			rcounts = RScounts(0)
			sele = "selected"
			if (StrComp(Rs("id"),request("id")))then
				sele = ""
			End If
	%>
                <option value="<%=Rs("id")%>" <%=sele%> ><%=Rs("categoria")%> - com <%=rcounts%> foto(s)</option>
	<%
		rs.MoveNext
	Loop 
rs.Close
Set rs = Nothing
	%>
              </select> </td>
          </tr>
          <tr> 
            <td><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">T&iacute;tulo:</font></div></td>
            <td><input  name="titulo" type="text" style="font-family: Arial; font-size: 8 pt;   WIDTH: 170px;" value="<%=titulo%>"></td>
          </tr>
          <tr> 
            <td><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Autor:</font></div></td>
            <td><input  name="autor" type="text" id="autor" style="font-family: Arial; font-size: 8 pt;  WIDTH: 170px;" value="agrofar"></td>
          </tr>
          <tr> 
            <td><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">E-Mail 
                Autor:</font></div></td>
            <td><input  name="mail" type="text" id="mail" style="font-family: Arial; font-size: 8 pt;   WIDTH: 170px;" value="agrofar@agrofar"></td>
          </tr>
          <tr> 
            <td><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o</font></div></td>
            <td><textarea name="descricao" rows="3" id="descricao" style="font-family: Arial; font-size: 8 pt;  WIDTH: 170px;"><%=descricao%></textarea></td>
          </tr>
          <tr> 
            <td><div align="center"> 
                <input type="button" value="Trocar Foto" onClick="location.href='fotos_upload.asp?img=deleta&temp_img=<%=foto%>&ac=<%=request("ac")%>&id=<%=request("id")%>'">
              </div></td>
            <td>&nbsp;</td>
            <td> <input type="submit"  value="<%=request("ac")%>"> </td>
          </tr>
        </table>

 </td></tr></table></form><p align="center"><!--#include file="rodape.asp"-->
