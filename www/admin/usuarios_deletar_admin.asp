
<!-- #include file="login_nivel3.asp" -->
<div align="center">
 
  <script  language="JavaScript">
//Funtion para selecionar todos os boxes de deletar
function checkAll(){

	for (i=0; i < document.frmDelete.chkDelete.length; i++){
		document.frmDelete.chkDelete[i].checked = document.frmDelete.chkAll.checked;
	}
}
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
</script>
  <%
	Sql2="DELETE FROM usuarios WHERE id=" & Request("ID")
	  Conn.Execute(Sql2)
	  'Sql3="DELETE FROM mail_interno WHERE id_para='" & Request("ID") & "'"
	  'Conn.Execute(Sql3)
	'  Sql4="DELETE FROM agenda WHERE id_user='" & Request("ID") & "'"
	 ' Conn.Execute(Sql4)
%>
  <font color="#FF0000" size="2" face="Arial, Helvetica, sans-serif"><strong>Este 
  usuário foi deletado com sucesso!<br>
  Dados contidos na sua agenda pessoal também foram apagados automaticamente.</strong></font><br>
  <script>alert('Dados delatados com sucesso!!!!!!  ');location.href='usuarios_default_admin.asp?p=False';</script>
  <%End If
%>
  <!--#include file="rodape.asp"-->
</div>

