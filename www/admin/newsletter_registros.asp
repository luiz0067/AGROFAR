 
<!-- #include file="login_nivel3.asp" -->
<!--#include file=".\config_site.inc"-->

<h2 align="center">
<hr color="#000000">
  <font size="3" face="Arial, Helvetica, sans-serif">Registro de E-mails enviados 
  da Lista de E-mails<br>
  </font>
  <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="newsletter_admin_menu.asp">Voltar 
  para página administradora de Newslettter</a></font> </h2></div>
<%
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Sqltext = "SELECT * FROM mail_list_send order by id DESC"
			Rs.Open SqlText, base_dados
			If Rs.EOF AND Rs.BOF = TRUE Then
			'nada				
		Else
		contador = 0
	Do Until Rs.EOF
	contador = contador + 1
								%>
<a href="newsletter_enviados_deletar.asp?id=<%=Rs("id")%>"><img src="../imagens/iFechar.gif" alt="Clique aqui para DELETAR" width="8" height="8" border="0"></a> 
<a href="javascript:mostrar_form_<%=contador%>();" title="Abre  a caixa de formulário para enviar perguntas" onMouseOver="MM_displayStatusMsg('Abre a caixa de formul&aacute;rio r&aacute;pido');return document.MM_returnValue"> 
<font size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=RS("data")%> - 
<%=RS("titulo")%> </font></a><br>
<script>
  <!--
    NS = eval("navigator.appName=='Netscape'")

	function apagar_form_<%=contador%>(){
		if (NS){
			document.layers.form_<%=contador%>.visibility = "hide"
		} else  {
			document.all.form_<%=contador%>.style.visibility = "hidden"
		}
	}
	
	function mostrar_form_<%=contador%>(){
    if (NS){
      document.layers.form_<%=contador%>.visibility = "show"
    } else  {
       document.all.form_<%=contador%>.style.visibility = "visible";
    }
  }
  //-->
</script>
<span id="form_<%=contador%>" name="form_<%=contador%>" style="visibility: hidden; position: absolute; left: 4px; top: 200px; width: 583px; z-index: 100; height: 228px;" height=220 width=300> 
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#999999" bordercolordark="#CCCCCC" bgcolor="#000000" >
  <tr> 
    <td width="100%" valign="middle" bgcolor="#CCCCCC" > <img src="../imagens/button_email.gif" width="21" height="20"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><%=Rs("titulo")%></b></font></td>
    <td width="0" height="0" bgcolor="#CCCCCC"><a href="javascript:apagar_form_<%=contador%>();"> <img src="../imagens/botaoX.gif" width="16" border=0 height="14"></a></td>
  </tr>
  <tr> 
    <td colspan="2" bgcolor="#FFFFFF"> <table width="100%" border="1" cellspacing="0" cellpadding="0">
        <tr> 
          <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
              <tr> 
                <td  bgcolor="#CCCCCC"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>De:<br>
                  </strong></font></td>
                <td bordercolor="#999999" bgcolor="#CCCCCC"><div align="left"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=site_titulo%></font></div></td>
              </tr>
              <tr> 
                <td bordercolor="#999999" bgcolor="#CCCCCC"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data:</strong></font></td>
                <td bordercolor="#999999" bgcolor="#CCCCCC"><div align="left"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=Rs("data")%></font></div></td>
              </tr>
              <tr> 
                <td bordercolor="#999999" bgcolor="#CCCCCC"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Para:</strong></font></td>
                <td bordercolor="#999999" bgcolor="#CCCCCC"><div align="left"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Enviado 
                    para <%=Rs("total_enviado")%> e-mails</font></div></td>
              </tr>
              <tr> 
                <td bordercolor="#999999" bgcolor="#CCCCCC"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Asunto:</strong></font></td>
                <td bordercolor="#999999" bgcolor="#CCCCCC"><div align="left"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=Rs("titulo")%></font></div></td>
              </tr>
              <tr> 
                <td colspan="2"> <%=Rs("texto")%></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td colspan="2" bordercolor="#999999" bgcolor="#CCCCCC"><acronym title=":: Clique aqui para fechar esta janela ::"> 
      <div align="right"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><strong><a href="javascript:apagar_form_<%=contador%>();"> 
        <img src="../imagens/iFechar.gif" alt=":: Clique aqui para fechar esta janela" width="7" height="7" border="0"><font size="1"> 
        Fechar janela</font></a> </strong></font> </div>
      </acronym></td>
  </tr>
</table>
</span>
<%
							Rs.MoveNext
							Loop
							End If
							
							Rs.Close
							set Rs = nothing
							%>
