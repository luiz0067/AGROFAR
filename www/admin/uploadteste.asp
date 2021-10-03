<!--#include file="login_nivel2.asp"-->
<%
if request("up") = "send" Then
	
	Set Upload = Server.CreateObject("SoftArtisans.FileUp") 
	Upload.Path = fotos_dir 
			
	Foto1 = Upload.Form("FILE1").ShortFileName 'campoFoto1 equivale ao nome do seu campo da foto1, no qual usamos 2 campos p/ fotos

	If Not Foto1 = "" Then
	  Upload.Form("FILE1").Save
	End If
			
        response.write " oi       a a a"& fotos_dir&Foto1 &"<br>"
End If	
        
%>
	<img src="<%=fotos_dir&Foto1%>">
	<FORM METHOD="POST" ENCTYPE="multipart/form-data" ACTION="?up=send&form=n&ac=<%=request("ac")%>&id=<%=request("id")%>">
		<input type="FILE" size="30" name="FILE1">
		<br>
		<input name="SUBMIT" type=SUBMIT class="f_vermelho_3" style="color: #ff0000; font-family: Arial;" value="Enviar Foto!">
	</FORM>
