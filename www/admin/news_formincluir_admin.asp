<!-- #include file="login_nivel1.asp" -->
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
<% IF Request.Form("mensagem") = "" THEN %>
<FORM METHOD="POST" NAME="XXX" ACTION="formincluir.asp?op=2" onSubmit="PegarTexto(this);">
  <div align="center">Titulo: 
    <input type="text" name="titulo">
    <br>
    Autor: 
    <input type="text" name="Autor">
    <br>
    E-mail : 
    <input type="text" name="email">
    <br>
    N&ordm; imagens: 
    <input type="text" name="aux" maxlength="1" size="1" value="0">
    <br>
	<TEXTAREA NAME="Mensagem" STYLE="display: none"></TEXTAREA>
    <OBJECT ID=myEditor DATA="editor.html" TYPE="text/x-scriptlet" WIDTH="500" HEIGHT="300">
      <!-- Isto só é exibido em browsers que não suportão scriptlets -->
    </OBJECT> <br>
    <INPUT TYPE="Submit" NAME="Submit" VALUE="Enviar">
    <INPUT TYPE="Reset" NAME="Reset" VALUE="Limpar o Formulário" OnClick="Limpar();">
  </div>
</FORM>
<% ELSE %>
   
<!--#include file=".\config_site.inc" -->
<%
   Set Conn = Server.CreateObject("ADODB.Connection")
   Conn.Open dominiundados
   Sql="INSERT INTO messages (Title,autor,mail,Message,online) VALUES('" & Request("titulo") & "','" & Request("autor") & "','" & Request("email") & "','" & Request("mensagem") & "',TRUE)"
   Conn.Execute(Sql)
   Conn.Close
   Set Conn = Nothing
%>
<% If Request.Form("aux") <> 0 Then %>  
  
<form name="form2" enctype="multipart/form-data" method="post" action="incluir_arquivos.asp?op=2">
  <% For I = 1 to Request.Form("aux") %>
    <input type="file" name="file<%= I %>">
    <br>
    <% Next %>
	<br>
	<input type="submit" name="Submit2" value="Upload Imagens">
    <br>
  </form>
<% Else %>
  <script> alert("Notícia postada com sucesso.") </script>
<%   Response.redirect "news_default_admin.asp" 
End If %>
<% END IF %>
</BODY>
