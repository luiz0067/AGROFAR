<!-- #include file="login_nivel3.asp" -->

  <!--#include file=".\config_site.inc"-->
 <div align="center"> :: ADMINISTRAÇAO DE MENUS ::<br>
<script>  
function del_menu_1(id){
	if (confirm("Tem ceteza que deseja deletar esse menu?\n Esse processo também irá deletar todos\n os itens relacionados a esse menu\n como também todas as páginas relacionadas.\n\n Opte por editar este ítem isto é mais seguro que deletar.")){
	pg="menu_deletar.asp?m=del&id_menu="+id;
	window.open(pg,'WinExa','titlebar=no,toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=0,width=240,height=140');
	 	
	} 
	else{
	return
	}
}
function del_menu_2(id,id2){
	if (confirm("Tem ceteza que deseja deletar esse ítem do menu,\n este processo apagar também a pagina de contúdo referente a este ítem.")){
	pg="menu_item_deletar.asp?it=del&id="+id+"&menu="+id2;
	window.open(pg,'WinExa','titlebar=no,toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=0,width=240,height=140');	
	} 
	else{
	return
	}
}
</script>
  
  <input name="button"  type="button" 
  style="cursor:hand;font-size:9pt;color: #CC0000; width:180px;"   
  onClick="window.open('menu_cadastrar.asp','WinExa','titlebar=no,toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=0,width=240,height=140');"  
  value="Cadastrar Novo Menu ">
  <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
  <input name="button32"  type="button" 
  style="cursor:hand;font-size:9pt;color: #CC0000; width:180px;"   
  onClick="window.open('menu_dinamico.asp','WinExa','titlebar=no,toolbar=0,location=0,status=0,menubar=0,scrollbars=1,resizable=0,width=100%,height=100%');"

  value="Ver como Esta o Menu" title="Clique Ver o Menu">
  </font></div>
  <%	Sqltext = "SELECT * FROM menu ORDER BY ordem ASC"
Set Rs = Conn.Execute(Sqltext)	%>

<table width="800"   border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center" >
    <tr> 
      <td> 
<table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr bgcolor="#3399FF"> 
          <td width="7%" background="../imagens/outros_icos/gallery/bg2.gif"><strong><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Ordem</font></strong></td>
          <td width="25%" background="../imagens/outros_icos/gallery/bg2.gif"><strong><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Menu</font></strong></td>
          <td width="43%" background="../imagens/outros_icos/gallery/bg2.gif"><strong><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o</font></strong></td>
          <td width="25%" background="../imagens/outros_icos/gallery/bg2.gif"><div align="center"><strong><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">A&ccedil;&atilde;o</font></strong></div></td>
        </tr>
        <% If Rs.EOF AND Rs.BOF = TRUE Then	%>
        <tr bgcolor="#FF3333"> 
          <td colspan="4" background="../imagens/forum_images/table_bg2.gif" bgcolor="#66CCFF">&nbsp;<br>
          Nenhum Menu Cadastrado!!!<br> </td>
        </tr>
        <%	Else	%>
        <%	Do While NOT RS.Eof  %>
        <tr bgcolor="#FF4848"> 
          <td background="../imagens/outros_icos/bd_vb2.gif"> <div align="center"><strong><font color="#000000" size="2" face="Arial, Helvetica, sans-serif"><%=Rs("ordem")%> 
              &nbsp;</font></strong></div></td>
          <td background="../imagens/outros_icos/bd_vb2.gif"> <li></li>
            <strong><font color="#000000" size="2" face="Arial, Helvetica, sans-serif">&nbsp;<%=Rs("menu")%></font></strong></td>
          <td background="../imagens/outros_icos/bd_vb2.gif"><strong><font color="#000000" size="2" face="Arial, Helvetica, sans-serif">&nbsp;<%=Rs("descri")%></font></strong></td>
          <td background="../imagens/outros_icos/bd_vb2.gif"> <div align="center"> 
              <input name="button"  type="button" 
  style="cursor:hand;font-size:8pt;color: #CC0000;width:80px;"   
  onClick="window.open('menu_editar.asp?id=<%=rs("id")%>','WinExa','titlebar=no,toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=0,width=240,height=140');"  
  value="Editar Menu" title="Clique aqui para Editar este Menu"><input name="button"  type="button" 
  style="cursor:hand;font-size:8pt;color: #CC0000; width:80px;"   
  onClick="del_menu_1('<%=rs("id")%>');"  
  value="Deletar Menu" title="Clique aqui para DELETAR este Menu. Cuidado, esta opao apaga junto todos os Sub-Menus, ou itens referentes a este Menu">
              <%	If Not rs("pg_ref")  Then 	%>
              <input name="button"  type="button"  title="Só cadastre uma Página de conteúdo para este Menu/Link, se o mesmo nao existir sub-itens de menu, caso o contrário opte por cadastrar páginas de conteúdo no ítens."
  style="cursor:hand;font-size:8pt;color: #CC0000; width:160px;"   
  onClick="location.href='menu_pg_form.asp?id_menu=<%=rs("id")%>&ac=cadastrar';"  
  value="Cadastrar Página de Conteúdo">
              <%	else	%>
              <input name="button"  type="button" title="Só Edite uma Página de conteúdo para este Menu/Link, se o mesmo nao existir sub-itens de menu, caso o contrário opte por cadastrar páginas de conteúdo no ítens. Caso já tenha castrado, ao cadastrar um Ítem de menu, esta página será cancelanda automaticamente por questoes de padrao de layout do site!!!"
  style="cursor:hand;font-size:8pt;color: #CC0000; width:160px;"   
  onClick="location.href='menu_pg_form.asp?id_menu=<%=rs("id")%>&ac=editar';"  
  value="Editar Página de Conteúdo">
              <%	End If	%>
            </div></td>
        </tr>
        <tr bgcolor="#990000" > 
          <td colspan="4" background="../imagens/forum_images/bg.gif"> 
            <!-- Entra tabela do itens do menu -->
            <table width="95%" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="6%" background="../imagens/forum_images/table_bg_image.gif" bgcolor="#990000"><font color="#CCCCCC" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Ordem</strong></font></td>
                <td width="31%" background="../imagens/forum_images/table_bg_image.gif" bgcolor="#990000"><strong><font color="#CCCCCC" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&Iacute;tens 
                  do Menu</font></strong></td>
                <td width="24%" background="../imagens/forum_images/table_bg_image.gif" bgcolor="#990000"><strong><font color="#CCCCCC" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
                  do &Iacute;tem</font></strong></td>
                <td width="39%" background="../imagens/forum_images/table_bg_image.gif" bgcolor="#990000"><font color="#FFFFFF"><strong><font color="#CCCCCC" size="1" face="Verdana, Arial, Helvetica, sans-serif">A&ccedil;&atilde;o 
                  em &iacute;tens de Menu</font></strong></font></td>
              </tr>
              <%
Sqltext2 = "SELECT * FROM menu_item WHERE id_menu = "&rs("id")&" ORDER BY ordem ASC"
Set Rs2 = Conn.Execute(Sqltext2)
If Rs2.EOF AND Rs2.BOF = TRUE Then
%>
              <tr > 
                <td colspan="4" bgcolor="#FFA4A4"background="../imagens/forum_images/table_bg.gif">&nbsp;<font size="2" face="Arial, Helvetica, sans-serif">Nenhum 
                  Item de menu cadastrado </font></td>
              </tr>
              <%	Else	%>
              <%	Do While NOT Rs2.EOF	%>
              <tr> 
                <td background="../imagens/forum_images/table_bg.gif" bgcolor="#FFA4A4"> 
                  <div align="center"> <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs2("ordem")%></font></div></td>
                <td background="../imagens/forum_images/table_bg.gif" bgcolor="#FFA4A4"> 
                  <Li></Li>
                  <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs2("menu_item")%></font></td>
                <td background="../imagens/forum_images/table_bg.gif" bgcolor="#FFA4A4"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs2("descri")%></font></td>
                <td background="../imagens/forum_images/table_bg.gif" bgcolor="#FFA4A4"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <input name="button3"  type="button" 
  style="cursor:hand;font-size:8pt;color: #0000ff; width:110px;"   
  onClick="window.open('menu_item_editar.asp?id=<%=rs2("id")%>&menu=<%=rs("menu")%>','WinExa','titlebar=no,toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=0,width=265,height=173');"

  value="Editar Ítem de Menu" title="Clique aqui para Editar este Menu"><input name="button3"  type="button" 
  style="cursor:hand;font-size:8pt;color: #0000ff; width:110px;"   
  onClick="del_menu_2('<%=rs2("id")%>','<%=rs("menu")%>');"  
  value="Deletar Ítem de Menu" title="Atençao, ao deletar este ítem desse Menu, estará deletando automaticamente a página de conteúdeo referente a este Menu!!!">
                    <%	If Not rs2("pg_ref")  Then 	%>
                    <input name="button4"  type="button" title="Ao Editar ou Cadastrar uma Página de Conteúdeo de um Ítem, automaticamente o sistema elimina a página de conteúdo do menu que origina estes ítens, isso aconte para manter o padrao do layout da página!!!"
  style="cursor:hand;font-size:8pt;color: #0000ff; width:220px;"   
  onClick="location.href='menu_pg_it_form.asp?id_menu_it=<%=rs2("id")%>&ac=cadastrar&id_menu=<%=rs2("id_menu")%>';"  
  value="Cadastrar Página Conteúdo Ítem">
                    <%	else	%>
                    <input name="button4"  type="button" title="Ao Editar ou Cadastrar uma Página de Conteúdeo de um Ítem, automaticamente o sistema elimina a página de conteúdo do menu que origina estes ítens, isso aconte para manter o padrao do layout da página!!!"
  style="cursor:hand;font-size:8pt;color: #0000ff; width:220px;"   
  onClick="location.href='menu_pg_it_form.asp?id_menu_it=<%=rs2("id")%>&ac=editar&id_menu=<%=rs2("id_menu")%>';"  
  value="Editar Página Conteúdo Ítem">
                    <%	End If	%>
                    </font> </div></font>                  </td>
              </tr>
              <tr>
                <td colspan="4"  bgcolor="#000000"></td>
              </tr>
              <%	Rs2.MoveNext
	Loop	
	Rs2.Close
	Set Rs2=nothing
	%>
              <%	End If %>
              <tr > 
                <td background="../imagens/forum_images/table_bg.gif"colspan="4" bgcolor="#FF9999"> 
                  <input name="button2"  type="button" 
  style="cursor:hand;font-size:8pt;color: #0000ff; "   
  onClick="window.open('menu_item_cadastrar.asp?id_menu=<%=rs("id")%>&menu=<%=rs("menu")%>','WinExa','titlebar=no,toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=0,width=265,height=173');"

  value="Cadastrar item do Menu <%=rs("menu")%>" title="Clique Cadastrar Ítem(s) do menu <%=rs("menu")%>"></td>
              </tr>
              <tr> 
                <td colspan="4">&nbsp;</td>
              </tr>
          </table></td>
          <%	Rs.MoveNext
	Loop	%>
          <%
Rs.Close
Set Rs=nothing
%>
          <%	End If	%>
        </tr>
      </table></td></tr>
</table>

<p><!-- #include file="rodape.asp" --></p>