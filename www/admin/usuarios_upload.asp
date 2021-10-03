<!--#include file=".\config_site.inc"-->



<title>:: Enviando sua Foto/Ícone::</title>
<body bgcolor="#CCCCCC" OnLoad="self.focus(); document.forms.frmImageUp.Submit.disabled=true;">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="0">
 <form action="usuarios_upload.asp?up=1" method="post" enctype="multipart/form-data" name="frmImageUp"  id="frmImageUp" onSubmit="alert('Por Favor aquarde imagem sendo enviada')">
  <tr> 
   <td width="100%"> 
    <fieldset>
        <legend><font size="2" face="Arial, Helvetica, sans-serif"><strong>Enviar 
        Imagem:</strong></font></legend>
        <table width="100%" border="0" cellspacing="0" cellpadding="1">
          <tr> 
            <td width="33%" align="right" ><font color="#FF0000" size="2" face="Arial, Helvetica, sans-serif">Caminho 
              da Foto</font><font size="2" face="Arial, Helvetica, sans-serif">:</font></td>
            <td width="82%"><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="file" type="file" size="30" class="text" onChange="document.forms.frmImageUp.Submit.disabled=false;" >
              </font></td>
          </tr>
          <tr align="center"> 
            <td colspan="2" ><font size="2" face="Arial, Helvetica, sans-serif">
			
<% 
'***********************************************************************
pers = true
dundas = false
'---------------------------

max =500000
max_b = max / 1000
caminho = usuario_dir
'Session("mensagem") = "3"

If request("up") = "1" Then

Set FSO = Server.Createobject("Scripting.FilesystemObject") 
if pers Then
Set Upload = Server.CreateObject("Persits.Upload")
	Upload.OverwriteFiles = False
	Upload.SetMaxSize 300000, True
	Upload.Save usuario_dir & "/original/"

Set FSO = Server.Createobject("Scripting.FilesystemObject") 
	For Each File in Upload.Files 
	
		If File.ImageType = "UNKNOWN" Then 'Se nao for uma imagem nao faz o upload
			Response.Write "O arquivo enviado não é uma imagem."
			File.Delete
			Response.End
		End If

	
	path = File.Path 
	Next 
	arquivo = FSO.GetFileName(path)
	
	
elseif dundas Then	
'on error resume next 
Set objUpload = Server.CreateObject("Dundas.Upload.2") 
objUpload.UseVirtualDir = false 
objUpload.Useuniquenames = true 
objUpload.MaxFileSize = 1048576 
objUpload.Save caminho & "/original/"

Set FSO = Server.Createobject("Scripting.FilesystemObject") 
For Each File in objUpload.Files 
path = File.Path 
Next 
arquivo = FSO.GetFileName(path)
end if



Set Jpeg = Server.CreateObject("Persits.Jpeg")
'path = caminho & "/" & arquivo  
Jpeg.Open path 
'FSO.Delete path

' qualidade da imagem
Interpolation =	1

if Jpeg.OriginalHeight > Jpeg.OriginalWidth Then
	tipo_img = "Imagem na vertical"
	Jpeg.Width = 60 'largura
	Jpeg.Height = 60 'altura
	Else
	tipo_img = "imagem na horizontal"
	Jpeg.Width = 60 'largura
	Jpeg.Height = 60 'altura
	End If
	
		original_largura = Jpeg.OriginalWidth
		original_altura = Jpeg.OriginalHeight
		pequena_largura = Jpeg.Width
		pequena_altura = Jpeg.Height
		
	Jpeg.Save caminho & "\pequena_" & arquivo
	
	
	'file.delete 
  Set Fil=Server.CreateObject("Scripting.FileSystemObject") 
	'Set filObject=Fil.GetFile 
	
	
			 Set filObject=Fil.GetFile(usuario_dir & "\original\" & arquivo) 
			 	filObject.Delete True
				arq = "upload/usuarios/pequena_" & arquivo
   %>
   
   
   <script  language="JavaScript">

	window.opener.config.foto.value = ('<%=arq%>');
	window.opener.config.avatar.src = ('../<%=arq%>');
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