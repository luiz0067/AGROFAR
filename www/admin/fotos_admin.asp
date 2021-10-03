<!--#include file="login_nivel2.asp"-->
<div align="center">
  <script language="JavaScript" type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
  <strong><font size="2" face="Arial, Helvetica, sans-serif">:: Administrador 
  da Galeria de Fotos :: </font></strong> </div>
<form name="form1"  >
  <table width="750" border="0" cellspacing="2" cellpadding="2" align="center">
  <tr> 
    <td><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">.: 
      Fotos ::..</font> </strong></td>
    <td>
        <select  name="menu1" class="f_vermelho_3_b"  onChange="MM_jumpMenu('self',this,0)">
          <option  selected >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;..:: Selecione 
          uma Categoria para ver ::..</option>
	<%
		  Set rs = Server.CreateObject("ADODB.Recordset")
	
	rs.Open "SELECT * FROM fotos_categorias order by id", conn

	Do Until rs.EOF
			sqlcount = "SELECT count(titulo) FROM fotos where id_cat=" & rs("id") 
			Set RScounts = Conn.Execute(sqlcount)
			rcounts = RScounts(0)
	%>
          <option value="?id_cat=<%=Rs("id")%>&amp;ver=categoria" ><%=Rs("categoria")%> - com <%=rcounts%> foto(s)</option>
          <%
rs.MoveNext
Loop 
rs.Close
Set rs = Nothing
	%>
        </select></td>
    <td>
    
    <table class="table">
      <tr>
        <td><div align="center"><a class="linque" href="javascript:location.href='fotos_categorias.asp'">Administrar Categorias</a></div></td>
      </tr>
    </table>
    
    </td>
    <td>
    
    <table class="table">
      <tr>
        <td><div align="center"><a class="linque" href="javascript:location.href='fotos_upload.asp?ac=Cadastrar&id=<%=request("id_cat")%>'">Cadastrar Nova Foto</a></div></td>
      </tr>
    </table>
    
    
    </td>
  </tr>
  <tr> 
    <td height="1" colspan="4" background="../imagens/icos_e_fundos/dh.gif"></td>
  </tr>
  <tr> 
    <td colspan="4"></td>
  </tr>
</table>
<!--#include file="fotos.asp"-->
<p align="center"><!--#include file="rodape.asp"-->