<!-- #include file="login_nivel2.asp" -->
<!--#include file="../enquete_incpollmentor.asp"-->
<%
Dim id
Dim count1, answer1
Dim count2, answer2
Dim count3, answer3
Dim count4, answer4
Dim count5, answer5
Dim count6, answer6
Dim count7, answer7
Dim count8, answer8


count1 = Request.Form("count1")
answer1 = Request.Form("answer1")
count2 = Request.Form("count2")
answer2 = Request.Form("answer2")
count3 = Request.Form("count3")
answer3 = Request.Form("answer3")
count4 = Request.Form("count4")
answer4 = Request.Form("answer4")
count5 = Request.Form("count5")
answer5 = Request.Form("answer5")
count6 = Request.Form("count6")
answer6 = Request.Form("answer6")
count7 = Request.Form("count7")
answer7 = Request.Form("answer7")
count8 = Request.Form("count8")
answer8 = Request.Form("answer8")

	Dim oConn
    Set oConn = PollMentor_GetDatabaseConn()
    Dim oRS
    
    
	If Request.QueryString("save")="yes" Then
		Set oRS = Server.CreateObject("ADODB.Recordset")
		
		Dim sID
		sID = Request.QueryString("id")
		If sID = "" Then
			sID = 0
		Else
			sID = CInt( sID)
		End If
		
		oRS.Open "select * from poll where id = " & sID, oConn, 1, 3
		
		If Request.QueryString("action") = "new" Or Request.QueryString("action") ="edit" Then
			If Request.QueryString("action") = "new"  Then
				oRS.AddNew
			End If
			If count1 = "" Then
				count1=0
			End If
			oRS("count1") = count1
			oRS("answer1") = answer1
			If count2 = "" Then
				count2=0
			End If
			oRS("count2") = count2
			oRS("answer2") = answer2
			If count3 = "" Then
				count3=0
			End If
			oRS("count3") = count3
			oRS("answer3") = answer3
			If count4 = "" Then
				count4=0
			End If
			oRS("count4") = count4
			oRS("answer4") = answer4
			If count5 = "" Then
				count5=0
			End If
			oRS("count5") = count5
			oRS("answer5") = answer5
			If count6 = "" Then
				count6=0
			End If
			oRS("count6") = count6
			oRS("answer6") = answer6
			If count7 = "" Then
				count7=0
			End If
			oRS("count7") = count7
			oRS("answer7") = answer7
			If count8 = "" Then
				count8=0
			End If
			oRS("count8") = count8
			oRS("answer8") = answer8
			oRS("question") = Request.Form("question")
		End If
		If Request.QueryString("action") = "del"  Then
				oRS.Delete
		Else
			oRS.Update
		End If
		oRS.Close
	Set oRS = Nothing
	oConn.Close
	Set oConn = Nothing
	%>
	<script>alert('Operação efetuada com sucesso !!!  ');location.href='enquete_default_admin.asp';</script>
	<%
	'Response.Redirect "admin_enquete_default.asp"
	End If
%>
<%
If Request.QueryString("action") = "edit" Then
	Dim oRS2
	Set oRS2 = oConn.Execute( "select * from poll where id = " & Request.QueryString("id") )
	question = oRS2("question")
	answer1 = oRS2("answer1")
	count1 = oRS2("count1")
	answer2 = oRS2("answer2")
	count2 = oRS2("count2")
	answer3 = oRS2("answer3")
	count3 = oRS2("count3")
	answer4 = oRS2("answer4")
	count4 = oRS2("count4")
	answer5 = oRS2("answer5")
	count5 = oRS2("count5")
	answer6 = oRS2("answer6")
	count6 = oRS2("count6")
	answer7 = oRS2("answer7")
	count7 = oRS2("count7")
	answer8 = oRS2("answer8")
	count8 = oRS2("count8")
	
	oRS2.Close
Else
End If
%>


<table width="90%"  border="0" align="center" cellPadding="3" cellSpacing="0" background="../imagens/forum_images/table_bg_image.gif" >
  <tbody>
    <tr background="../imagens/forum_images/table_bg_image.gif"> 
      <td width="50%" height="35" vAlign="top" background="../imagens/forum_images/table_bg_image.gif" >&nbsp; 
      </td>
      <td width="468" height="35" vAlign="top" background="../imagens/forum_images/table_bg_image.gif" > 
        <p align="center"> <b><font color="#FFFFFF" face="verdana,arial,helvetica" size="+2">Administrador 
          de Enquêtes</font></b> </p>
      </td>
    </tr>
    <tr>
      <td height="" vAlign="top" width="100%" colspan="2" background="../imagens/forum_images/table_bg_image.gif"> 
        <table align="center" bgColor="#ffffff" border="0" cellPadding="0" cellSpacing="0" height="100%" width="100%">
          <tbody>
            <tr>
              <td height="100%" vAlign="top" width="85%">
                <table bgColor="#ffffff" border="0" cellPadding="10" cellSpacing="0" height="100%" width="100%" ../imagens/forum_images/table_bg_image.gif>
                  <tbody>
                    <tr>
                      <td align="left" height="100%" vAlign="top" width="65%">
                        <hr color="#000066" noShade SIZE="1">
                             
                             
                        <table border="0" width="100%">
                          <tr>
                            <td width="70%">
                             
                        <p align="left"><font color="#FF9900"><b><font color="#000000" size="2">Enquête</font></b><br>
                        <%
                        Dim sURL
                        sURL = "enquete_poll_admin.asp?save=yes&action=" & Request.QueryString("action")
                       	sURL = sURL & "&id=" & Request.QueryString("id") 
                        %>
						</font>
                        <form method="POST" action="<%=sURL%>">
                                <table border="0" width="100%" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td colspan="2" bgcolor="#FFFFFF"><font color="#003366" size="2" face="Arial, Helvetica, sans-serif"><b>Pergunta: 
                                      <input type="text" name="question" size="66" value="<%=question%>" >
                                      </b></font></td>
                                  </tr>
                                  <tr> 
                                    <td width="29%"><font color="#003366" size="2" face="Arial, Helvetica, sans-serif"><b>Resposta 
                                      1:<br>
                                      <input type="text" name="answer1" size="33" value="<%=answer1%>" >
                                      </b></font></td>
                                    <td width="71%"><font color="#003366" size="2" face="Arial, Helvetica, sans-serif"><b>Nº 
                                      de votos<br>
                                      <input type="text" name="count1" size="20" value="<%=count1%>" >
                                      </b></font></td>
                                  </tr>
                                  <tr> 
                                    <td width="29%"><font color="#003366" size="2" face="Arial, Helvetica, sans-serif"><b>Resposta&nbsp; 
                                      2:<br>
                                      <input type="text" name="answer2" size="33" value="<%=answer2%>" >
                                      </b></font></td>
                                    <td width="71%"><font color="#003366" size="2" face="Arial, Helvetica, sans-serif"><b>Nº 
                                      de votos:<br>
                                      <input type="text" name="count2" size="20" value="<%=count2%>" >
                                      </b></font></td>
                                  </tr>
                                  <tr> 
                                    <td width="29%"><font color="#003366" size="2" face="Arial, Helvetica, sans-serif"><b>Resposta&nbsp; 
                                      3:<br>
                                      <input type="text" name="answer3" size="33" value="<%=answer3%>" >
                                      </b></font></td>
                                    <td width="71%"><font color="#003366" size="2" face="Arial, Helvetica, sans-serif"><b>Nº 
                                      de votos:<br>
                                      <input type="text" name="count3" size="20" value="<%=count3%>" >
                                      </b></font></td>
                                  </tr>
                                  <tr> 
                                    <td width="29%"><font color="#003366" size="2" face="Arial, Helvetica, sans-serif"><b>Resposta&nbsp; 
                                      4:<br>
                                      <input type="text" name="answer4" size="33" value="<%=answer4%>" >
                                      </b></font></td>
                                    <td width="71%"><font color="#003366" size="2" face="Arial, Helvetica, sans-serif"><b>Nº 
                                      de votos:<br>
                                      <input type="text" name="count4" size="20" value="<%=count4%>" >
                                      </b></font></td>
                                  </tr>
                                  <tr> 
                                    <td width="29%"><font color="#003366" size="2" face="Arial, Helvetica, sans-serif"><b>Resposta&nbsp; 
                                      5:<br>
                                      <input type="text" name="answer5" size="33" value="<%=answer5%>" >
                                      </b></font></td>
                                    <td width="71%"><font color="#003366" size="2" face="Arial, Helvetica, sans-serif"><b>Nº 
                                      de votos:<br>
                                      <input type="text" name="count5" size="20" value="<%=count5%>" >
                                      </b></font></td>
                                  </tr>
                                  <tr> 
                                    <td width="29%"><font color="#003366" size="2" face="Arial, Helvetica, sans-serif"><b>Resposta 
                                      6:<br>
                                      <input type="text" name="answer6" size="33" value="<%=answer6%>" >
                                      </b></font></td>
                                    <td width="71%"><font color="#003366" size="2" face="Arial, Helvetica, sans-serif"><b>Nº 
                                      de votos:<br>
                                      <input type="text" name="count6" size="20" value="<%=count6%>" >
                                      </b></font></td>
                                  </tr>
                                  <tr> 
                                    <td width="29%"><font color="#003366" size="2" face="Arial, Helvetica, sans-serif"><b>Resposta 
                                      7:<br>
                                      <input type="text" name="answer7" size="33" value="<%=answer7%>" >
                                      </b></font></td>
                                    <td width="71%"><font color="#003366" size="2" face="Arial, Helvetica, sans-serif"><b>Nº 
                                      de votos:<br>
                                      <input type="text" name="count7" size="20" value="<%=count7%>" >
                                      </b></font></td>
                                  </tr>
                                  <tr> 
                                    <td width="29%"><font color="#003366" size="2" face="Arial, Helvetica, sans-serif"><b>Resposta 
                                      8:<br>
                                      <input type="text" name="answer8" size="33" value="<%=answer8%>" >
                                      </b></font></td>
                                    <td width="71%"><font color="#003366" size="2" face="Arial, Helvetica, sans-serif"><b>Nº 
                                      de votos:<br>
                                      <input type="text" name="count8" size="20" value="<%=count8%>" >
                                      </b></font></td>
                                  </tr>
                                </table>
                              
                                  <input type="submit" value="Criar Enquête" name="B1" >
                                 
                        </form>
                            
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

</body>

</html>
