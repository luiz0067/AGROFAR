<!-- #include file="login_nivel2.asp" -->
<!--#include file="../enquete_incpollmentor.asp"-->

<%


    Set oConn = PollMentor_GetDatabaseConn()
    
	Dim strSort
	strSort = Request.QueryString("sort")    
	If strSort = "" Then
		strSort = " active, createdwhen desc"
	Else
		strSort = strSort & " desc"
	End If
	
	If Request.QueryString("saveactive") = "yes" Then
		Dim sActive
		Response.Write Request.Form("ACTIVE")
		oConn.Execute "update poll set active=false where active=true"
		oConn.Execute "update poll set active=true where id=" & Request.Form("ACTIVE")
	End If
	
%>




<table align="center" bgColor="#003399" border="0" cellPadding="3" cellSpacing="0" height="233" width="90%" background="../imagens/forum_images/table_bg_image.gif">
  <tbody>
    <tr>
      <td width="50%" height="37" vAlign="top" background="../imagens/forum_images/table_bg_image.gif" >&nbsp; 
      </td>
      <td width="468" height="37" vAlign="top" background="../imagens/forum_images/table_bg_image.gif" > 
        <b><font color="#FFFFFF" face="verdana,arial,helvetica" size="+2">Administrador 
        de Enquêtes</font></b> </td>
    </tr>
    <tr>
      <td height="185" vAlign="top" width="100%" colspan="2" background="../imagens/forum_images/table_bg_image.gif">
        <table align="center" bgColor="#ffffff" border="0" cellPadding="0" cellSpacing="0" height="80%" width="100%">
          <tbody>
            <tr>
              <td height="80%" vAlign="top" width="85%">
                <table bgColor="#ffffff" border="0" cellPadding="10" cellSpacing="0" height="100%" width="100%">
                  <tbody>
                    <tr>
                      <td align="left" height="100%" vAlign="top" width="65%">
                        <font color="#aa3333" face="verdana,arial,helvetica" size="+2">
                        <hr color="#000066" noShade SIZE="1">
                             
                        </font>
                             
                        <table border="0" width="100%">
                          <tr>
                            <td width="70%"> <font face="helvetica, arial" size="2"><b><i><font face="Arial, Helvetica, sans-serif">Essas 
                              são as enquêtes que atualmente estão no seu banco 
                              de dados.&nbsp;<br>
                              Escolha as opções para administrar seu sistema de 
                              enquête.&nbsp;</font></i></b></font> <form method="POST" action="enquete_default_admin.asp?saveactive=yes">
                                <table border="0" width="100%" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td background="../imagens/forum_images/table_bg.gif" bgcolor="<%=cor_tabela%>"><font size="2" face="Arial, Helvetica, sans-serif"><b><font color="#000000">Ativar</font></b></font></td>
                                    <td background="../imagens/forum_images/table_bg.gif" bgcolor="<%=cor_tabela%>"><font size="2" face="Arial, Helvetica, sans-serif"><b><a href="enquete_poll_admin.asp?sort=question"><font color="#000000">Pergunta</font></a></b></font></td>
                                    <td background="../imagens/forum_images/table_bg.gif" bgcolor="<%=cor_tabela%>"><font size="2" face="Arial, Helvetica, sans-serif"><b><a href="enquete_default_admin.asp?sort=createdwhen"><font color="#000000">Adicionado 
                                      em</font></a></b></font></td>
                                    <td width="100" background="../imagens/forum_images/table_bg.gif" bgcolor="<%=cor_tabela%>"><div align="center"><font size="2" face="Arial, Helvetica, sans-serif"><b><font color="#000000">Menu</font></b></font></div></td>
                                  </tr>
                                  <%
	Dim oRS
	Set oRS = oConn.Execute("select * from poll order by " & strSort)
	Dim bgcolor
	bgcolor = "#ECECD9"
	while not oRS.EOF
%>
                                  <tr> 
                                    <%
                          Dim sBold, sBoldStop, sSel
                          If oRS("active") = True Then
                          	sBold = "<b>"
                          	sBoldStop ="</b>"
                          	sSel = "checked"
                          Else
                             sBold = ""
                             sBoldStop = ""
                             sSel = ""
                          End If
                          %>
                                    <td bgcolor="<%=bgcolor%>"> <p><font size="2" face="Arial, Helvetica, sans-serif">
                                        <input type="radio" value="<%=oRS("id")%>" <%=sSel%> name="ACTIVE">
                                        <%=sBold%></font></td>
                                    <td bgcolor="<%=bgcolor%>"><font size="2" face="Arial, Helvetica, sans-serif"><%=sBold%><%=Trim(oRS("question"))%><%=sBoldStop%></font></td>
                                    <td bgcolor="<%=bgcolor%>"><div align="center"><font size="2" face="Arial, Helvetica, sans-serif"><%=sBold%><%=oRS("createdwhen")%><%=sBoldStop%></font></div></td>
                                    <td width="100" bgcolor="<%=bgcolor%>"><font size="2" face="Arial, Helvetica, sans-serif"><b><a href="enquete_poll_admin.asp?id=<%=oRS("id")%>&action=edit">Editar</a> 
                                      - <a href="enquete_poll_admin.asp?id=<%=oRS("id")%>&save=yes&action=del">Deletar</a></b></font></td>
                                  </tr>
                                  <%	
	If bgcolor="#ECECD9" Then
		bgcolor = "#FFFFFF"
	Else
		bgcolor="#ECECD9"
	End if
	oRS.MoveNext
	Wend
%>
                                </table>
                                <font size="2" face="Arial, Helvetica, sans-serif"><br>
                                <input type="submit" value="Ativar Enquête!!" name="B1" >
                                </font>
                              </form>
                              <p><font size="2" face="Arial, Helvetica, sans-serif"><a href="enquete_poll_admin.asp?action=new"><b>..:: 
                                Adicionar Nova Enquête ::..</b></a></font></p>
                             
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </tbody>
                </table>
              </td>
            </tr>
          </tbody>
        </table>
      </td>
    </tr>
  </tbody>
</table>
<!--#include file="rodape.asp"-->
</body>

</html>
