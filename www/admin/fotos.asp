


<script language="JavaScript" type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>


<style type="text/css">
<!--
.style2 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 8pt;
	color: #E80000;
}
.style3 {
color: #000000;
font-family: Arial, Helvetica, sans-serif;
	font-size: 7pt;
	}
.style4 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 8pt;
	color: #333333;
}
.style8 {font-family: Arial, Helvetica, sans-serif; font-size: 8pt; color: #333333; }
-->
</style>
<script>
function ver_com(id){
	if (document.getElementById('com_v_'+id).style.display == "none") {
		document.getElementById('com_v_'+id).style.display = "";
		
			} else {
		document.getElementById('com_v_'+id).style.display = "none";
		
		}
}
function com_envia(id,tit,cat){
	//if (document.getElementById('com_e').style.display == "none") {
		document.getElementById('com_e').style.display = "";
		form2.autor.focus();
		form2.id_foto.value = id;
		form2.tit.value = tit;
		form2.id_cat.value = cat;
		//	} else {
		//document.getElementById('com_e').style.display = "none";
		
		//}
}
</script>



<script type="text/javascript" src="../css/amplia.js"></script>
<script type="text/javascript">var path='../upload/fotos/'; var xoffset = -1;</script>


<DIV class=layer_fotos id=container 
                        style="DISPLAY: block; Z-INDEX: 1; VISIBILITY: hidden; POSITION: absolute; left: 59px; top: 361px;"> 
  <TABLE cellSpacing=0 cellPadding=5 border=0 onClick="close();">
                          <TBODY>
                          <TR>
                            
        <TD class=institucional align=right height=20><A 
                              href="javascript:fechar();"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">FECHAR 
          [X]</font> </A></TD>
      </TR>
                          <TR>
                            
        <TD align=middle vAlign=center bgcolor="#F5F5F5"><a href="javascript:fechar();"><IMG 
                              class=img_border id=picture alt="Clique na Foto para Fechar" 
                              src="<%=logo_empresa%>" ></a></TD>
      </TR>
                          <TR>
                            <TD vAlign=center align=middle>
          <DIV class=legenda 
                        id=legenda></DIV>          </TD></TR></TBODY></TABLE></DIV>

   
    <%	if request("come") = "deleta" Then	
	
	Sql2="DELETE FROM fotos_comentarios WHERE id=" & Request("ID") 
  Conn.Execute(Sql2)
  
	%>
<table width="350" border="0" cellspacing="1" cellpadding="1" align="center" class="moldura1">
  <tr>
    <td height="31"><div align="center"><img src="../imagens/icos_e_fundos/warning.png" align="absmiddle"> <span style="color: #FF0000; font-weight: bold"> O Comentário foi DELETADO com Sucesso!</span></div>
      <div align="center"><br />
        O comentário enviado por <span class="style2"><%=request("autor")%></span> foi deletado.</div><br /></td>
  </tr>
</table>

    <%	End If	%>
    
     <%	if request("foto") = "deleta" Then	
	
	Sql2="DELETE FROM fotos WHERE id=" & Request("ID") 
  Conn.Execute(Sql2)
  Sql3="DELETE FROM fotos_comentarios WHERE id_foto='" & Request("ID") & "'"
  Conn.Execute(Sql3)
  
	%>
<table width="350" border="0" cellspacing="1" cellpadding="1" align="center" class="moldura1">
  <tr>
    <td height="31"><div align="center"><img src="../imagens/icos_e_fundos/warning.png" align="absmiddle"> <span style="color: #FF0000; font-weight: bold"> O Foto foi DELETADA com Sucesso!</span></div>
      <div align="center"><br />
       A foto <span class="style2"><%=request("titulo")%></span> foi deletada e todos<br /> 
       os comentários da mesma também.</div>
      <br /></td>
  </tr>
</table>

    <%	End If	%>
<%
  
		  Set rs = Server.CreateObject("ADODB.Recordset")
if request("ver") = "categoria" Then
	rs.Open "SELECT  * FROM fotos where id_cat = "&request("id_cat")&" order by id DESC", conn
	Else
	rs.Open "SELECT Top 20 * FROM fotos order by id DESC", conn
	End If
	%>
<table width="660" border="0" align="center" cellpadding="4" cellspacing="4">
	 <tr>
	 <% 
	 	tr = 0
		ct = 0
	 Do Until rs.EOF
	 tr = tr + 1
If tr = 4 Then
Response.Write("<Tr>")
End If
	%>
 <%
				  conta = "SELECT count(id) as total from fotos_comentarios WHERE id_foto='"&rs("id")&"'" 
				Set Rs3 = Conn.Execute(conta)
				total_com = Rs3("total")
				Rs3.Close
				Set RS3=nothing
				  %>
  
     <td valign="top" >
     
     <table width="190" border="0" align="center" cellpadding="4" cellspacing="4">
  <tr>
    <td height="190" valign="bottom"  class="moldura2" ><div align="center">
    <A href="javascript:abre('g_<%=rs("foto")%>.jpg','<%=rs("titulo")%>')">
    <img src="../upload/fotos/p_<%=rs("foto")%>.jpg"  border="0" title="Clique na foto para ampliar">
    </A>
    <br />
     <span class="style3">clique na foto para ampliar<br /></span></div>
     <table class="table2">
<tr>
  <td><div align="center"><a class="linque" href="fotos_editar_admin.asp?ac=Editar&id=<%=Rs("id")%>&temp_img=<%=Rs("foto")%>">Editar Foto</a></div></td>
  <td><div align="center"><a class="linque" href="?foto=deleta&id_cat=<%=Rs("id_cat")%>&ver=categoria&titulo=<%=rs("titulo")%>&id=<%=rs("id")%>" 
  OnClick="return confirm('Você tem certeza que você quer excluir esta Foto?\n também irá excluir todos os comentários dessa foto.\nEsta ação não poderá ser desfeita!!!')">Deletar Foto</a></div></td>
</tr>
</table>
  
     
     </td>
  </tr>
  <tr>
    <td class="moldura1"><div align="center"><span class="style2">"<%=rs("titulo")%>"</span></div></td>
  </tr>
  <tr>
    <td class="moldura1"> <div align="center"><span class="style8">&copy; <%=rs("autor")%>-<%=rs("data")%><br />
     
    </span></div>
       
        <div align="left"><span class="style8">
          <a href="javascript:ver_com('<%=rs("id")%>');" class="style8"><img src="../imagens/icos_e_fundos/write.png" alt="Clique para ver os comentários enviados nesta foto" width="16" height="16" align="absmiddle" border="0">ver e deletar comentários (<%=total_com%>)</a><br>
        </span></div></td>
  </tr>
  <tr>
    <td class="moldura1" style="display:none;" id="com_v_<%=rs("id")%>">
    <%
						
Sqltext = "SELECT * FROM fotos_comentarios WHERE id_foto = '"& rs("id")&"' ORDER BY id DESC"
						Set Rs2 = Conn.Execute(Sqltext)
						If Rs2.EOF AND Rs2.BOF = TRUE Then
						%>
                          <div align="center"><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Nenhum 
                            Comentário cadastrado.<br>
                            </font></div>
                          
                          <%	Else	%>
            <%	Do Until Rs2.EOF	%>
                          <span class="style2"><a href="?come=deleta&id=<%=Rs2("id")%>&id_cat=<%=Rs("id_cat")%>&ver=categoria&autor=<%=rs2("nome")%>"><img src="../imagens/icones/delete.png" alt="DELETAR" width="16" height="16" border="0" align="absmiddle" 
                          
                          OnClick="return confirm('Você tem certeza que você quer excluir este Comentário?\n Esta ação não poderá ser desfeita!!!\n')"/></a> <%=rs2("comentario")%><br />
            </span><span class="style8">Enviado por <a href="mailto:<%=rs2("mail")%>"><%=rs2("nome")%></a></span>
                          <br /><center><img src="../imagens/icos_e_fundos/pt.gif" /></center>
                           <%	Rs2.MoveNext
							Loop
							End If
						%>    </td>
  </tr>
  </table>

     </td>
     <td >&nbsp;</td>
     
   
    <%
	If tr = 3 Then
Response.Write("</Tr>")

tr = 0
End If

rs.MoveNext
Loop 
rs.Close
Set rs = Nothing
	%>
</table>
