<!--#include file="nivel.asp"-->
<HEAD>

<SCRIPT LANGUAGE="JavaScript">
  function PegarTexto(f) {
    f.elements.Mensagem.value = document.all.myEditor.html + "";
  }
  function Limpar() {
    document.all.myEditor.html = "";
  }
</SCRIPT>

</HEAD>
<BODY>
<div align="center">
<!--#include file="news_menu_admin.asp" -->
</div>
<!--#include file=".\config_site.inc" -->
<%
   Set Conn1 = Server.CreateObject("ADODB.Connection")
   Conn1.Open dominiundados
%>
<% IF request.form("cat") = "" THEN %>
<FORM METHOD="POST" NAME="XXX" ACTION="formenviarnews.asp?op=3" onSubmit="PegarTexto(this);">
  <div align="center">
    <select size="1" name="cat" >
          <option selected value="0">Selecione uma Categoria</option>
          <%
		      SQL = "SELECT codigo, descr FROM categoriamail ORDER BY descr"
              Set RS = Conn1.Execute(SQL)
		  %>
		  <% WHILE NOT RS.Eof 
		               response.write "<option  value=""" &  RS("codigo") &""">" & RS("descr") & "</option>"
					   RS.MoveNext
		        WEND %>
	 </select>
    <br>
	<TEXTAREA NAME="Mensagem" STYLE="display: none"></TEXTAREA>
    <OBJECT ID=myEditor DATA="newseditor.asp" TYPE="text/x-scriptlet" WIDTH="500" HEIGHT="300">
      <!-- Isto só é exibido em browsers que não suportão scriptlets -->
    </OBJECT> <br>
    <INPUT TYPE="Submit" NAME="Submit" VALUE="Enviar">
    <INPUT TYPE="Reset" NAME="Reset" VALUE="Limpar o Formulário" OnClick="Limpar();">
  </div>
</FORM>
<% ELSE
   IF Request.form("cat") = 0 THEN
      SQL = "SELECT * FROM newsletter WHERE ativo=true"
   ELSE
      SQL = "SELECT * FROM newsletter WHERE op=" & request.form("cat") & " and ativo=true"
   END IF
   Set RS = Conn1.Execute(SQL)
   While Not RS.Eof
      para = para & RS("email") & "; "
	  Rs.movenext
   Wend
str = str + "<html>"
str = str + "<head>"
str = str + "<title>:: Dominiun Montanhismo ::</title>"
str = str + "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
str = str + "</head>"
str = str + "<body bgcolor=""#FFFFFF"" text=""#000000"">"
str = str + Request.Form("mensagem")
str = str + "</body>"
str = str + "</html> "              
			    set oMail = Server.CreateObject("CDONTS.NewMail")

				'str = request.form("texto") & vbcrlf & vbcrlf & "Nome: "
				'str = str & request.form("nome") & " - E-mail: "
				'str = str & request.form("email")				
				
				oMail.To=para
				oMail.From="luiz0067@yahoo.com.br"
				oMail.Subject="Newsletter - Dominiun Montanhismo"
				oMail.Body = str
				oMail.Importance=2
				oMail.BodyFormat=0
				oMail.MailFormat=0
				oMail.Send
				set oMail=Nothing
   END IF %>

</BODY>
 <%
      RS.Close
      Conn1.Close
      Set RS = Nothing
      Set Conn1 = Nothing
    %>


