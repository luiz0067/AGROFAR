<!--#include file="login_nivel2.asp"-->
<div align="center" ><strong><font size="2" face="Arial, Helvetica, sans-serif">:: 
  Administrando Categorias de Fotos:: </font></strong><br>
  <br>
  <script language="JavaScript" type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>


  <table width="530" border="0" cellspacing="3" cellpadding="3">
  <tr>
    <td>
    <!-- -->
    <table class="table">
      <tr>
        <td><div align="center"><a class="linque" href="fotos_categorias.asp?form=cadastrar">Cadastrar Categorias</a></div></td>
      </tr>
    </table>
    <!-- -->    </td>
    <td><!-- -->
    <table class="table">
      <tr>
        <td><div align="center"><a class="linque" href="javascript:com_edita();">Editar Categorias</a></div></td>
      </tr>
    </table>
    <!-- --></td>
    <td><!-- -->
    <table class="table">
      <tr>
        <td><div align="center"><a class="linque" href="javascript:com_deleta();">Deletar Categorias</a></div></td>
      </tr>
    </table>
    <!-- --></td>
  </tr>
  <tr>
    <td colspan="3" id="com_edita"  style="display:none;"  class="moldura1"><form>
      <div align="center"> <br />
        <select  name="menu1" class="f_vermelho_3_b" style="color: #000000; font-family: Arial; font-size: 9 pt;" onChange="MM_jumpMenu('self',this,0)">
                <option selected>Selecione 
                  uma Categoria para EDITAR</option>
          
          <%
		  Set rs = Server.CreateObject("ADODB.Recordset")
	
	rs.Open "SELECT * FROM fotos_categorias where categoria<>'Patrocinadores' order by id", conn

	Do Until rs.EOF
	sqlcount = "SELECT count(titulo) FROM fotos where id_cat=" & rs("id") 
			Set RScounts = Conn.Execute(sqlcount)
			rcounts = RScounts(0)
			
	%>
          <option value="?id=<%=Rs("id")%>&form=editar&cat_nome=<%=Rs("categoria")%>" ><%=Rs("categoria")%> - com <%=rcounts%> foto(s)</option>
          <%
rs.MoveNext
Loop 
rs.Close
Set rs = Nothing
	%>
        </select>
      </div>
    </form></td>
    </tr>
  <tr>
    <td colspan="3" style="display:none;" id="com_deleta" class="moldura1"><form>
      <div align="center"> <br />
      <select  name="menu1" class="f_vermelho_3_b" style="color: #000000; font-family: Arial; font-size: 9 pt;" onChange="MM_jumpMenu('self',this,0);">
        <option selected>Selecione 
          uma Categoria para DELETAR</option>
        
        <%
		  Set rs = Server.CreateObject("ADODB.Recordset")
	
	rs.Open "SELECT * FROM fotos_categorias where categoria<>'Patrocinadores' order by id", conn

	Do Until rs.EOF
	sqlcount = "SELECT count(titulo) FROM fotos where id_cat=" & rs("id") 
			Set RScounts = Conn.Execute(sqlcount)
			rcounts = RScounts(0)
			if rcounts = 0 then
	%>
        <option value="?id=<%=Rs("id")%>&form=deletar&cat_nome=<%=Rs("categoria")%>" ><%=Rs("categoria")%> - com <%=rcounts%> foto(s)</option>
        <%
		End If
rs.MoveNext
Loop 
rs.Close
Set rs = Nothing
	%>
        </select>
      <br /></div>
      <span class="f_preto_4"><img src="../imagens/icos_e_fundos/warning.png" alt="ATEN&Ccedil;&Atilde;O!!!" width="16" height="16" align="absmiddle" /> S&oacute; &eacute; possivel deletar categorias com nenhuma foto cadastrado.<br />
      <img src="../imagens/icos_e_fundos/warning.png" alt="ATEN&Ccedil;&Atilde;O!!!" width="16" height="16" align="absmiddle" /> Caso voc&ecirc; queira deletar uma categoria que tenha fotos, primeiro v&aacute; em Editar Foto e muda a sua categoria, e somente depois volte para esta sess&atilde;o para deletar.</span>
    </form></td>
    </tr>
</table>
<script>

function com_edita(){
		if (document.getElementById('com_deleta').style.display == "") {
		document.getElementById('com_deleta').style.display = "none";
		}
		document.getElementById('com_edita').style.display = "";
}
function com_deleta(){
		if (document.getElementById('com_edita').style.display == "") {
		document.getElementById('com_edita').style.display = "none";
		}
		document.getElementById('com_deleta').style.display = "";
}
</script>
</div>
<%	if request("form") = "cadastrar" Then %>
<FORM METHOD="POST" ACTION="?cat=cadastrar" > 
<table width="300" border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center" >
                    <tr> 
                      <td>
  
        
        
  <table  width="300"border="0" align="center" cellpadding="3" cellspacing="0" >
          <tr bgcolor="#F3F3F3"> 
            <td colspan="3" align="right" background="../imagens/icos_e_fundos/table_bg.gif"> 
              <div align="center"><font color="#FFFFFF"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Cadastrar 
            Categoria de Fotos</font></strong></font> </div></td>
          </tr>
          <tr bgcolor="#F3F3F3"> 
            <td width="100%" align="right"><strong><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><span class="titulo" id="var1">Categoria:&nbsp;</span></font></strong></td>
            <td colspan="2" nowrap> <strong><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"> 
              <input name="categoria" type="text"   size="35">
              </font> </strong></td>
          </tr>
          
          <tr bgcolor="#F3F3F3"> 
            <td width="100%" align="right">&nbsp;</td>
            <td colspan="2" valign="top"><font color="#006600" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <input name="Submit" type="submit" class="f_vermelho_3" style=" color: #000000" value="Cadastrar">
              </font></td>
          </tr>
        </table>    
 </td></tr></table>
</form>
<% end if
	if request("form") = "editar" Then %>
<FORM METHOD="POST" ACTION="?cat=editar&id=<%=request("id")%>" > 
<table width="300" border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center" >
                    <tr> 
                      <td>
  
        
        
  <table  width="300" border="0" align="center" cellpadding="3" cellspacing="0" >
          <tr bgcolor="#F3F3F3"> 
            <td colspan="3" align="right" background="../imagens/icos_e_fundos/table_bg.gif"> 
              <div align="center"><font color="#FFFFFF"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Editar  
            Categoria de Fotos</font></strong></font> </div></td>
          </tr>
          <tr bgcolor="#F3F3F3"> 
            <td width="100%" align="right"><strong><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><span class="titulo" id="var1">Categoria:&nbsp;</span></font></strong></td>
            <td colspan="2" nowrap> <strong><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"> 
              <input name="categoria" type="text"  value="<%=request("cat_nome")%>"  size="35">
              </font> </strong></td>
          </tr>
          
          <tr bgcolor="#F3F3F3"> 
            <td width="100%" align="right">&nbsp;</td>
            <td colspan="2" valign="top"><font color="#006600" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <input name="Submit" type="submit" class="f_vermelho_3" style=" color: #000000" value="Editar">
              </font></td>
          </tr>
        </table>    
 </td></tr></table>
</form>
<%	End If
	  if request("cat") = "cadastrar" Then
	  
	 Sql = "INSERT INTO fotos_categorias (categoria)"
	 sql = sql + " VALUES('" & Request.Form("categoria") & "')"
		Conn.Execute(Sql)
		
	  	%>
		<script>
	   	msg = " CATEGORIA CADASTRADA \n";
		msg += "________________________\n";
		msg += "A categoria:<%=request("categoria")%>\n";
		msg += "cadastrada com Sucesso!!!\n";
		alert(msg);
		location.href="fotos_admin.asp";
	   </script>
	   <%	End If
	    if request("cat") = "editar" Then
		 Sql="UPDATE  fotos_categorias SET categoria = '" & Request("categoria") &  "' WHERE id =" & request("id")     		
	       Conn.Execute(Sql)
	   	%>
			<script>
	   	msg = " CATEGORIA EDITADA \n";
		msg += "________________________\n";
		msg += "A categoria:<%=request("categoria")%>\n";
		msg += "Editada com Sucesso!!!\n";
		alert(msg);
		location.href="fotos_admin.asp";
	   </script>
	   <%	End If
       
	    if request("form") = "deletar" Then
		 Sql="DELETE FROM  fotos_categorias WHERE id =" & request("id")     		
	       Conn.Execute(Sql)
	   	%>
			<script>
	   	msg = " CATEGORIA DELETADA \n";
		msg += "________________________\n";
		msg += "A categoria:<%=request("cat_nome")%>\n";
		msg += "Editada com Sucesso!!!\n";
		alert(msg);
		location.href="fotos_admin.asp";
	   </script>
	   <%	End If	%>
       <br />
       <img src="../imagens/icos_e_fundos/previous.png" alt="VOLTAR" align="absmiddle" /><a href="fotos_admin.asp">Voltar para Fotos
       </a><br />
<br />
	  <p align="center"><!--#include file="rodape.asp"-->