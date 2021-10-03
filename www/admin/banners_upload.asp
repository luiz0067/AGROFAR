<%
Response.buffer = true
%>
<!-- #include file="login_nivel3.asp" -->
<!--#include file=".\config_site.inc" -->
	
<html>
   <head>
   <title>Upload de Arquivos</title>
   </head>
   <body>
<%

	IF Request.QueryString("acao") = "form" Then
%> <FORM METHOD="POST" ENCTYPE="multipart/form-data" ACTION="banners_upload.asp?acao=upload&B_ID=<%=Request.QueryString("B_ID")%>">
	<table>
   <tr valign="baseline"> 
           
            <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000080">IMAGEM
              URL:</font></b></td>
            <td> 
              <input type="file" name="B_IMAGE" value="" size="60">
              <input type="Submit" name="Submit2" value="Upload da Imagem" onclick="">
            </td>
          </tr>
          </table>
<% 
end if


IF Request.QueryString("acao") = "upload" Then

'caminho = "E:\escalada_brasil\upload\croquis\"	
	caminho = "c:\domains\dominiun.com\wwwroot\upload\banners\" 

Set objUpload = Server.CreateObject("Dundas.Upload.2") 
objUpload.UseVirtualDir = false 
objUpload.Useuniquenames = true 
'objUpload.MaxFileSize = 1048576 


objUpload.Save caminho
Set FSO = Server.Createobject("Scripting.FilesystemObject") 


For Each File in objUpload.Files 
path = File.Path 
Next 
arquivo = FSO.GetFileName(path)




   
   				novo_arquivo = "upload/banners/"  & arquivo 	 
	 
	 	  		Sql = "UPDATE banners SET B_IMAGE='" & novo_arquivo & "' WHERE B_ID=" & Request.QueryString("B_ID")
  								Conn.Execute(Sql)

    
  								'Sql = "UPDATE banners SET B_IMAGE='" & arquivo & "' WHERE B_ID=" & Request.QueryString("B_ID")
  								'Conn.Execute(Sql)
  
  'RS.Close
  'Set RS = Nothing
  

  %>
   
   Arquivo Enviado

<% 
response.redirect "banner_admin.asp" 


 end if

%>
 