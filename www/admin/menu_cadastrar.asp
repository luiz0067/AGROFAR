<!-- #include file="login_nivel3_limpo.asp" -->
<!-- #include file=".\config_site.inc" -->
<title>Cadastrar Menu do Site &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</title>
<base target="_self">
<form name="form1" method="post" action="?ac=cadastrar">
  <table width="200" border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center">
    <tr> 
      <td> <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
          <tr > 
            <td height="23" colspan="5" background="../imagens/forum_images/table_bg.gif"> 
              <center>
                <strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">:: 
                CADASTRAR MENUS::</font> </strong> </center></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td width="35%" bgcolor="#EEEEEE"> <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;Menu: 
                </font></div></td>
            <td colspan="2"> 
              <input name="menu" type="text" size="33" >
             </td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td width="35%" height="24" bgcolor="#EEEEEE"> <div align="right" ><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o:</font></div></td>
            <td colspan="2">
              <input name="descri" type="text" size="33" >
            </td>
          </tr>
          
          <tr bgcolor="#EEEEEE"> 
            <td colspan="3" bgcolor="#EEEEEE"> <div align="center"><font color="#000000" size="2">&nbsp;</font><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ordem 
                de Apresenta&ccedil;&atilde;o:</font><font size="1" face="Arial, Helvetica, sans-serif"> 
                <select name="ordem">
                  <option>1</option>
                  <option>2</option>
                  <option>3</option>
                  <option>4</option>
                  <option>5</option>
                  <option>6</option>
                  <option>7</option>
                  <option>8</option>
                  <option>9</option>
                  <option>10</option>
                  <option>11</option>
                  <option>12</option>
                </select>
                </font></div></td>
          </tr>
          <tr bgcolor="#EEEEEE"> 
            <td bgcolor="#EEEEEE"> <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"> 
                </font></div></td>
            <td colspan="2" bgcolor="#EEEEEE"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp; 
              <input name="imageField" type="image" src="../imagens/outros_icos/botao_cadastrar.gif" width="129" height="24" border="0">
              </font></td>
          </tr>
        </table></td>
    </tr>
  </table>
 
  </form>
<% IF REQUEST("ac")= "cadastrar" then
Sql="INSERT INTO menu (menu,descri,ordem) VALUES('" & Request("menu") & "','" & Request("descri") & "'," & Request("ordem") & ")"
  Conn.Execute(Sql)
%>
Menu Cadastrado com Sucesso!!!<br>
<br>
<script>
window.opener.location.href="menus_default.asp";
self.close();
//alter("Menu Cadastrado com Sucesso!!");location.href="";</script>

<% End If	%>
