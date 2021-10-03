<!--#include file=".\config_site.inc"-->
<LINK 
href="http://www.unicred.com.br/imagens/unicred.ico" rel="SHORTCUT ICON">
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1"><LINK 
href="functions/estilos.css" rel=stylesheet><LINK 
href="functions/links_styles.css" rel=stylesheet>

<script>
//Script de Menu By Midiamix

//----------------------------------
function submenu(menu){
sm = document.getElementById(menu);

if(sm.style.display=="none"){
	ul = document.getElementById("menus").getElementsByTagName("div");
for(i=0;i<ul.length;i++){
ul.item(i).style.display="none";
}
sm.style.display="block";
} else {
sm.style.display="none";
}
}
</script>

<%	Sqltext = "SELECT * FROM menu ORDER BY ordem ASC"
Set Rs = Conn.Execute(Sqltext)	%>
<% If Rs.EOF AND Rs.BOF = TRUE Then	%>
<body background="../imagens/uni_img/m/m_r2_c2.gif">
Nenhum Menu Cadastrado!!!<br>
<%	Else	%>
<div id="menus">

<table width="70%" border="0" cellspacing="1" cellpadding="1" background="../imagens/uni_img/m/m_r2_c2.gif">
  
<%	Do While NOT RS.Eof  %>
<tr>
<td onClick="submenu('<%=Rs("id")%>');" style="cursor:hand" title="<%=Rs("descri")%>" class=linkmenu><a href="#" class="linkmenu"><%=Rs("menu")%></a></td>
</tr> 
<%
Sqltext2 = "SELECT * FROM menu_item WHERE id_menu = "&rs("id")&" ORDER BY ordem ASC"
Set Rs2 = Conn.Execute(Sqltext2)
If Rs2.EOF AND Rs2.BOF = TRUE Then
%>
<%	Else	%><tr><td>
<div id="<%=Rs("id")%>" style="display:none">
<table>
<%	Do While NOT Rs2.EOF	%>
<tr >
              <td><img src="../imagens/img_fundo/seta_num.gif"><a href="#" class="linksub"><%=rs2("menu_item")%></a></td>
  </tr> 
<%	Rs2.MoveNext
	Loop	
	Rs2.Close
	Set Rs2=nothing
	%>
</table>	</div></td></tr>
<%	End If %>
<tr>
    <td bgcolor="#CCCC66"></td>
  </tr>
   
<%	Rs.MoveNext
	Loop	%>
<%
Rs.Close
Set Rs=nothing
%>
<%	End If	%>

  <tr>
      <td  ><a href="#" class="linkmenu">Notícias</a></td>
  </tr> 
  <tr>
    <td bgcolor="#CCCC66"></td>
  </tr> 
  <tr>
    <td ><a href="#" class="linkmenu">Contato</a></td>
  </tr>  
  
</table>
</div>
