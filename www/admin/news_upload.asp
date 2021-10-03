<!--#include file=".\config_site.inc"-->



<title>Imagens Upload</title>
<body bgcolor="#CCCCCC" OnLoad="self.focus(); document.forms.frmImageUp.Submit.disabled=true;">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="0">
 <form action="news_upload.asp?upload=enviar" method="post" enctype="multipart/form-data" name="frmImageUp"  id="frmImageUp" onSubmit="alert('Por Favor aquarde imagem sendo enviada')">
  <tr> 
   <td width="100%"> 
    <fieldset>
        <legend><font size="2" face="Arial, Helvetica, sans-serif"><strong>Enviar 
        Imagem:</strong></font></legend>
        <table width="100%" border="0" cellspacing="0" cellpadding="1">
          <tr> 
            <td width="10%" align="right" class="text"><font size="2" face="Arial, Helvetica, sans-serif">Procurar 
              imagem:<br>
              Legenda da Foto:</font></td>
            <td width="90%"><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="file" type="file" size="35" onChange="document.forms.frmImageUp.Submit.disabled=false;" >
              <br>
              <input name="legenda" type="text" size="55" id="legenda">
              </font></td>
          </tr>
          <tr align="center"> 
            <td colspan="2" class="text"><font size="2" face="Arial, Helvetica, sans-serif">
			
<% 
salva = noticias_dir 'Server.MapPath(".")
tamanho = 7000000 
pode = tamanho / 1024 

If request("upload") = "enviar" Then

Set Upload = Server.CreateObject("Persits.Upload")
Upload.CodePage = 65001
Upload.SetMaxSize tamanho, True '
t = 7000000 / 1024
On Error Resume Next

Set Jpeg = Server.CreateObject("Persits.Jpeg")
Upload.Save salva 'Server.MapPath(".")
If Err.Number = 8 Then	%>
<font size="2" face="Verdana, Arial, Helvetica, sans-serif">O arquivo é muito 
Pesado.<br>
Desculpe mas o arquivo que você está tentando enviar é muito pesado, o <strong>máximo</strong> 
permitido é de<strong> 280Kb</strong>.<br>
Por favor tente novamente!!!<br>
Visite a sessão de ajuda do nosso site para saber como alterar o tamanho do seu 
croqui.</font> <br>
<%	Else
For Each File in Upload.Files
	If File.ImageType <> "JPG" Then
		Response.Write "Desculpe mas este não é um arquivo JPG."
		File.Delete
		Response.End
	End If
leg = Server.HTMLEncode(Item.Value)
response.Write("legenda:"& Server.HTMLEncode(Item.Value))
	'Jpeg.Open File.Path
		
randomize timer	
	gera = (9876989856 * CInt((RND * 32000) + 100 ))
Jpeg.Open File.Path

' qualidade da imagem
'Interpolation =	1
'Jpeg.Canvas.Font.Color = &006600' red
'Jpeg.Canvas.Font.BkMode = "Transparent"
'Jpeg.Canvas.Font.Family = "ARIAL"
'Jpeg.Canvas.Font.Size = 7
'Jpeg.Canvas.Font.Bold = True
'Jpeg.Canvas.Print 2, 2, "®" & site_titulo
if Jpeg.OriginalHeight > Jpeg.OriginalWidth Then
	tipo_img = "Imagem na vertical"
	Jpeg.Width = 160 'largura
	Jpeg.Height = 200 'altura
	Else
	tipo_img = "imagem na horizontal"
	Jpeg.Width = 200 'largura
	Jpeg.Height = 160 'altura
	End If
	
		original_largura = Jpeg.OriginalWidth
		original_altura = Jpeg.OriginalHeight
		pequena_largura = Jpeg.Width
		pequena_altura = Jpeg.Height
		Jpeg.Save salva & "\p_"&gera&".jpg" 'Imagem 800x600

if Jpeg.OriginalHeight > Jpeg.OriginalWidth Then
	tipo_img = "Imagem na vertical"
	Jpeg.Width = 56 'largura
	Jpeg.Height = 70 'altura
	Else
	tipo_img = "imagem na horizontal"
	Jpeg.Width = 70 'largura
	Jpeg.Height = 56 'altura
	End If
		Jpeg.Save salva & "\mini_"&gera&".jpg" 'Imagem 800x600
		
		
		
		Jpeg.Open File.Path
	
if Jpeg.OriginalHeight > Jpeg.OriginalWidth Then
		tipo_img = "Imagem na vertical"
		Jpeg.Width = 400 'largura
		Jpeg.Height = 600 'altura
	Else
		tipo_img = "imagem na horizontal"
		Jpeg.Width = 600 'largura
		Jpeg.Height = 400 'altura
	End If
		original_largura = Jpeg.OriginalWidth
		original_altura = Jpeg.OriginalHeight
		pequena_largura = Jpeg.Width
		pequena_altura = Jpeg.Height
		
		Jpeg.Save salva & "\g_"&gera&".jpg" 'Imagem 800x600

		
	file.delete 

 
     
For Each Item in Upload.Form
	legenda = Server.HTMLEncode(Item.Value)
Next
%>  
   legenda:<%=legenda%>
   <script>
   window.opener.frames.message.document.focus();
	//window.opener.frames.message.document.execCommand("InsertImage", false, '<%=site_end%>/upload/noticias/p_<%=gera%>.jpg"' + ' OnClick=open(g_<%=gera%>.jpg,);'); //   OnClick=open(g_' + '<%=gera%>.jpg , ' + '<%=leg%>);');
	var t 
	//t = "<%=legenda%>";
	t = frmImageUp.legenda.value; 
	img1 = "<%=site_end%>/upload/noticias/p_<%=gera%>.jpg"
	img2 = "g_<%=gera%>.jpg"
	
	//window.opener.frames.message.document.body.innerHTML =  window.opener.frames.message.document.body.innerHTML + '<a href="javascript:open(|_aspas_|g_<%=gera%>.jpg|_aspas_|,|_aspas_|<%=legenda%>|_aspas_|)"> <IMG src="<%=site_end%>/upload/noticias/p_<%=gera%>.jpg"  border="0" alt="Clique na Imagem para Ampliar" style="border: 1px solid #000000;margin: 5px;"></a>' ;
	 window.opener.frames.message.document.body.innerHTML =  window.opener.frames.message.document.body.innerHTML + '<center><div class="moldura2" > <a href="javascript:abre(|_aspas_|g_<%=gera%>.jpg|_aspas_|,|_aspas_|<%=legenda%>|_aspas_|)"> <IMG src="<%=site_end%>/upload/noticias/p_<%=gera%>.jpg"  border="0" alt="Clique na Imagem para Ampliar" ><br><center><img src="<%=site_end%>/imagens/icones/zoom_mini.gif" alt="Ampliar Foto" align="absmiddle" border=0> </a> <%=legenda%></center></div></center><br>' ;
	 
	window.close();
	</script>
              <%
'End If
'Next


Set objUpload = Nothing
Set Jpeg = Nothing


Next
end if
eND IF
%>            </td>
          </tr>
        </table>
    </fieldset></td>
  </tr>
  <tr align="right"> 
   <td> <input type="submit"  value="     Enviar    "> &nbsp; <input type="button" name="cancel" value=" Cancelar " onClick="window.close()"></td>
  </tr>
 </form>
</table>
</body></html>