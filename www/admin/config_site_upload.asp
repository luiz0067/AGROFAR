<!--#include file=".\config_site.inc"-->



<title>:: Enviando sua logomarca ::</title>
<body bgcolor="#CCCCCC" OnLoad="self.focus(); document.forms.frmImageUp.Submit.disabled=true;">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="0">
 <form action="config_site_upload.asp?up=1" method="post" enctype="multipart/form-data" name="frmImageUp"  id="frmImageUp" onSubmit="alert('Por Favor aquarde imagem sendo enviada')">
  <tr> 
   <td width="100%"> 
    <fieldset>
        <legend><font size="2" face="Arial, Helvetica, sans-serif"><strong>Enviar 
        Imagem:</strong></font></legend>
        <table width="100%" border="0" cellspacing="0" cellpadding="1">
          <tr> 
            <td width="33%" align="right" ><font color="#FF0000" size="2" face="Arial, Helvetica, sans-serif">Caminho 
              da Logo</font><font size="2" face="Arial, Helvetica, sans-serif">:</font></td>
            <td width="82%"><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="file" type="file" size="30" class="text" onChange="document.forms.frmImageUp.Submit.disabled=false;" >
              </font></td>
          </tr>
          <tr align="center"> 
            <td colspan="2" ><font size="2" face="Arial, Helvetica, sans-serif">
			
<% 
'***********************************************************************
If request("up") = "1" Then

	Set Upload = Server.CreateObject("Persits.Upload")
	Upload.OverwriteFiles = False
	Upload.SetMaxSize 300000, True
	Upload.Save logo_dir 

	For Each File in Upload.Files

		novo_arquivo = Upload.Files(1).ExtractFileName
	
		If File.ImageType = "UNKNOWN" Then 'Se nao for uma imagem nao faz o upload
			Response.Write "O arquivo enviado não é uma imagem."
			File.Delete
			Response.End
		End If
	Next
response.Redirect "?up=2&arquivo=" & novo_arquivo

End If

If request("up") = "2" Then

	Set Jpeg = Server.CreateObject("Persits.Jpeg")
	Jpeg.New 308, 110, &HFFFFFF
	'Jpeg.Canvas.Pen.Color = False '&H000080 ' Blue COLOCAR COR DE BORDA
	'Jpeg.Canvas.Brush.Solid = False ' or a solid bar would be drawn
	'Jpeg.Canvas.DrawBar 1, 1, Jpeg.Width, Jpeg.Height

	Set Img = Server.CreateObject("Persits.Jpeg")
		
	Img.Open logo_dir & "/" & Request("arquivo")  
			
			' Resize to inscribe in 100x100 square
			Img.PreserveAspectRatio = True
			if Img.OriginalHeight > Img.OriginalWidth Then 'imagem vertical
				Img.Width = 100
				Img.Height = 100
			Else 'Imagem horizontal
			Img.Width = 180
			Img.Height = 100
			End If
			X = 90
			Y = 5

				Jpeg.Canvas.DrawImage X + (100 - Img.Width)/2, Y + (100 - Img.Height)/2, Img
		
	arq = "n_" & request("arquivo")
	Jpeg.Save 	logo_dir & "/" & arq  
'End If
	Scale = 50
	Jpeg.Width = Jpeg.OriginalWidth * Scale / 100
	Jpeg.Height = Jpeg.OriginalHeight * Scale / 100
	
	Jpeg.Save logo_dir & "\m_" & arq
	Set Fil=Server.CreateObject("Scripting.FileSystemObject") 

	Set filObject=Fil.GetFile(logo_dir & "\" & request("arquivo")) 
	filObject.Delete True	
   %>
   
   
   <script>
   window.opener.config.logo_empresa.focus();
	window.opener.config.logo_empresa.value=('upload/' + '<%=arq%>');
	//window.opener.config.logo_model.value=('upload/' + '<%=arq%>');
	
	//img_obj = form.image_file.value;
    //document.changing.src = img_obj;
    //document.changing2.src= img_obj;
    //document.changing.width = document.changing2.width;
    //document.changing.height = document.changing2.height;
	//src
	window.close();
	</script>
              <%
'End If
'Next
Set objUpload = Nothing
Set Jpeg = Nothing
Set Img = Nothing

End If
if Session("mensagem") = "1" Then
Session("mensagem") = "O arquivo que você está tentando enviar não uma imagem JPG, só permitimos enviar imagens no formato JPG."
End If
if Session("mensagem") = "2" Then
Session("mensagem") = "O arquivo que você está tentando enviar exede o máximo permitido que é de " & max_b & "kb"
End If
if Session("mensagem") = "3" Then
Session("mensagem") = "O máximo permitido para enviar arquivos é de " & max_b & "kb"
End If
	%>
	<%=Session("mensagem")%>
	</font>
            </td>
          </tr>
        </table>
    </fieldset></td>
  </tr>
  <tr align="right"> 
   <td> <input type="submit" name="Submit" value="     Enviar    "> &nbsp; <input type="button" name="cancel" value=" Cancelar " onClick="window.close()"></td>
  </tr>
 </form>
</table>
</body></html>