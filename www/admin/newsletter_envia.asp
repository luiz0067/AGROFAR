
<%

Server.ScriptTimeout = 21600
%>
<!--#include file="../newsletter_mail_func.asp"-->

<div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Enviando 
  e-mails</strong></font> 

</div>


<div align="center"><br>
  <font size="1" face="Verdana, Arial, Helvetica, sans-serif">Ps-Não clique em 
  REFRESH ou ATUALIZAR, se não alguns membros receberão e-mail duas vezes!<br>
  <br>
  Isto pode demorar algum tempo dependendo da velocidade do servidor de correio 
  e quantos e-mails há para enviar.</font><br>
  <form name="frmSent">
  <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Enviando 
    <input type="text" size="4" name="sent" value="0">
    de um total de e-mails para serem enviados.</strong></font> </div>
   </form>
  <%	
Session("strMailBody") = Request("body")

	Sql = "SELECT * FROM mail_list WHERE online = true"
	   	Set news = Conn.Execute(Sql)
		if news.EOF AND news.BOF =True Then
		' nenhum mail para enviar
Else
 
 If Request.Form("Submit") = "Enviar um Preview para você" Then
 enviar_teste
 response.Redirect("newsletter_form.asp?send_voce=ok&form=ver")
 Else
 Session("strMailBody") = ""
 contador = 0

Do Until news.EOF
session("email") = news("email")
contador = contador + 1

Response.Write(vbCrLf & "<script langauge=""JavaScript"">document.frmSent.sent.value = " & contador & ";</script>")
enviar_todos


   news.MoveNext
   Loop
   Session("email") = ""
   'gravando na lista de enviados
   	newsletter = "SELECT count(id) as total_send from mail_list_send" 
	Set news = Conn.Execute(newsletter)
	
IF news.EOF THEN
total_send = 0
else
total_send = news("total_send") 
total_send = total_send + 1

		SQL = "INSERT INTO mail_list_send (mail_n,total_enviado,titulo,texto,data)"
		SQL = SQL + "VALUES ('" & total_send & "','" & contador & "','" & site_titulo & " Newsletter Nº:" & total_send & "','" & Request.form("body") & "',"
		SQL = SQL + "'" & Day(Now) & ("/") & month(Now) & ("/") & Year(Now) & "')"
		
		Conn.Execute(Sql) 
		End if
		News.close
		set news=nothing
   %>
  <br>
  <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Seu email 
  foi enviado agora a todos os sócios de sua lista de e-mails</strong></font>. 
  <%
   Set Mail = Nothing
   End If
   End If

   %>
</div>
<p></p>

