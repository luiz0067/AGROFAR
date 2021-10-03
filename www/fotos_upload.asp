<!--#include file="login_nivel2.asp"-->
<% 
salva = fotos_dir 'Server.MapPath(".")
tamanho = 307200 
pode = tamanho / 1024 

if request("img") = "deleta" Then
on error resume next
	Set Fil=Server.CreateObject("Scripting.FileSystemObject") 
		Set filObject=Fil.GetFile(fotos_dir & "\m_" & request("temp_img")&".jpg") 
			 	filObject.Delete True
			Set filObject=Fil.GetFile(fotos_dir & "\p_" & request("temp_img")&".jpg") 
			 	filObject.Delete True
			Set filObject=Fil.GetFile(fotos_dir & "\g_" & request("temp_img")&".jpg") 
			 	filObject.Delete True
	
End If

if not request("form") = "n" then	%>
<FORM METHOD="POST" ENCTYPE="multipart/form-data" ACTION="?up=send&form=n&ac=<%=request("ac")%>&id=<%=request("id")%>">
  <table width="350" border="0" cellspacing="0" cellpadding="1" bgcolor="#CCCCCC" align="center" >
                          <tr> 
                            <td> 
							
							<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
            <td height="23" background="../imagens/icos_e_fundos/table_bg2.gif"> 
              <div align="center" class="f_vermelho_3_b">Selecione 
            a Foto/Imagem para enviar:</div></td>
  </tr>
  <tr>
            <td align="center" bgcolor="#F0F0F0"> <font size="1" face="Arial, Helvetica, sans-serif"><img src="../imagens/icos_e_fundos/info.png" width="16" height="16" align="absmiddle"> 
              O Sistema s&oacute; aceita croquis enviados em JPG.</font><br> <font size="1" face="Arial, Helvetica, sans-serif"><img src="../imagens/icones/picture.png" width="16" height="16" align="absmiddle"> 
        O tamanho m&aacute;ximo do arquivo enviado &eacute; de <%=Formatnumber(fix(tamanho/1024),0)%> kb.</font><br> 
        <input type="FILE" size="30" name="FILE1">
              <br>
              <input name="SUBMIT" type=SUBMIT class="f_vermelho_3" style="color: #ff0000; font-family: Arial;" value="Enviar Foto!">
              <br>
            </td>
  </tr>
</table>
</td></tr></table>
  </FORM>


<%
End If
if request("up") = "send" Then

'Set Upload = Server.CreateObject("Persits.Upload")
Set Upload = Server.CreateObject("SoftArtisans.FileUp") 

Upload.Path = logo_dir

If Upload.IsEmpty Then
		response.write "<center>Por favor, indique um arquivo para upload.</center><br>"
Else
	'Salva o arquivo no servidor
	Upload.Save
	response.write "<center>Total de Bytes Enviados: " & Upload.TotalBytes & "</center>"
End if

'Gera um link html para retornar a pagina anterior
response.write "<center><a href='javascript:history.go(-1)'>Voltar</a></center>"

Set SaFileUp = Nothing





Set Jpeg = Server.CreateObject("Persits.Jpeg")

' Capture and save uploaded image to the same directory as script
'Upload.Save salva 'Server.MapPath(".")

For Each File in Upload.Files
	'ae_nome = File.Name
	'ae_end = File.Path 
	'ae_tamanho = File.Size 
	'exte = File.Ext
	 
	if File.Size > 307200 Then
	%>
    <table width="650" border="0" cellspacing="1" cellpadding="1" align="center" class="moldura1">
  <tr>
    <td height="21"><p align="center"><img src="../imagens/icos_e_fundos/warning.png" align="absmiddle"> <span class="f_vermelho_3_b" style="color: #FF0000; font-weight: bold">ERRO AO ENVIAR ARQUIVO!</span>
          </p>
      <p><span class="f_vermelho_3_b">O arquivo é muito 
        Pesado.</span><br>
        Desculpe mas o arquivo que voce está tentando enviar é muito pesado, o <strong>máximo</strong> permitido é de<strong> 300 kb</strong>.<br>
        O arquivo que voce está enviando tem (<strong><%=Formatnumber(fix(file.size/1024),0)%> kb</strong>).<br />
        Por favor tente novamente!!!<br>
        <br>
        <img src="../imagens/icos_e_fundos/question.png" alt="Ajuda" width="16" height="16" align="absmiddle"> Esta com <span style="font-weight: bold">problemas</span> para reduzir o <span style="font-weight: bold">tamanho</span> da sua <span style="font-weight: bold">foto</span>, ou entao quer <span style="font-weight: bold">converte-la</span> em formato <span style="font-weight: bold">JPG</span>?<br> 
        <img src="../imagens/icos_e_fundos/question.png" alt="Ajuda" width="16" height="16" align="absmiddle"> Visite o site <a href="http://media-convert.com" target="_blank" class="f_vermelho_1_b">http://media-convert.com</a> e converta seus arquivos online de maneira rápida e simples. <br>
        <img src="../imagens/icos_e_fundos/question.png" alt="Ajuda" width="16" height="16" align="absmiddle"> O tamanho que usamos é 
        de <span style="font-weight: bold">largura</span> (Width) = <span style="font-weight: bold">480</span>, <span style="font-weight: bold">altura</span> (Height) = <span style="font-weight: bold">640</span> para imagens na <span style="font-weight: bold">vertical</span>, para <span style="font-weight: bold">horizontal</span> <span style="font-weight: bold">largura</span> (Width) = <span style="font-weight: bold">640</span>, altura (Height) = <span style="font-weight: bold">480</span><br />
        <br>
        
        <a href="javascript:history.go(-1);"><img src="../imagens/icos_e_fundos/previous.png" alt="Voltar" border="0"  align="absmiddle"/>VOLTAR</a></font>
        <br>
        <br>
          </p></td>
  </tr>
</table>
<!--#include file="rodape.asp"-->
<%
		File.Delete
		Response.End

	ElseIf File.ImageType <> "JPG" Then
	%>
  <table width="650" border="0" cellspacing="1" cellpadding="1" align="center" class="moldura1">
  <tr>
    <td height="21"><p align="center"><img src="../imagens/icos_e_fundos/warning.png" align="absmiddle"> <span class="f_vermelho_3_b" style="color: #FF0000; font-weight: bold">ERRO AO ENVIAR ARQUIVO!</span>
          </p> <p><span class="f_vermelho_3_b">Formato/Extensao do arquivo nao permitido.</span><br>
    Desculpe mas o arquivo que voce está enviando nao é um arquivo JPEG (<strong>.jpg</strong>) é um arquivo (<strong><%=File.Ext%></strong>).<br />
    Por questao de segurança, nosso sistema só permite o envio de arquivos neste formato.<br />
    <br>
 <br>
        <img src="../imagens/icos_e_fundos/question.png" alt="Ajuda" width="16" height="16" align="absmiddle"> Esta com <span style="font-weight: bold">problemas</span> para reduzir o <span style="font-weight: bold">tamanho</span> da sua <span style="font-weight: bold">foto</span>, ou entao quer <span style="font-weight: bold">converte-la</span> em formato <span style="font-weight: bold">JPG</span>?<br> 
        <img src="../imagens/icos_e_fundos/question.png" alt="Ajuda" width="16" height="16" align="absmiddle"> Visite o site <a href="http://media-convert.com" target="_blank" class="f_vermelho_1_b">http://media-convert.com</a> e converta seus arquivos online de maneira rápida e simples. <br>
        <img src="../imagens/icos_e_fundos/question.png" alt="Ajuda" width="16" height="16" align="absmiddle"> O tamanho que usamos é 
        de <span style="font-weight: bold">largura</span> (Width) = <span style="font-weight: bold">480</span>, <span style="font-weight: bold">altura</span> (Height) = <span style="font-weight: bold">640</span> para imagens na <span style="font-weight: bold">vertical</span>, para <span style="font-weight: bold">horizontal</span> <span style="font-weight: bold">largura</span> (Width) = <span style="font-weight: bold">640</span>, altura (Height) = <span style="font-weight: bold">480</span><br />
        <br>
        
<a href="javascript:history.go(-1);"><img src="../imagens/icos_e_fundos/previous.png" alt="Voltar" border="0"  align="absmiddle"/>VOLTAR</a></font><br><br>
 </p></td>
  </tr>
</table>
<!--#include file="rodape.asp"-->
<%
		File.Delete
		Response.End
	Else

randomize timer	
	gera = (9876989856 * CInt((RND * 32000) + 100 ))

Jpeg.Open File.Path

	Interpolation =	1
	Jpeg.Canvas.Font.Color = &000000' red
	Jpeg.Canvas.Font.BkMode = "Transparent"
	Jpeg.Canvas.Font.Family = "ARIAL"
	Jpeg.Canvas.Font.Size = 12
'	jpeg.Canvas.Font.Opacity = 0.9
	'Jpeg.Canvas.Font.Bold = True
	Jpeg.Canvas.Print 2, 2, "®"&year(now)&" "&site_titulo

if Jpeg.OriginalHeight > Jpeg.OriginalWidth Then
		tipo_img = "Imagem na vertical"
		Jpeg.Width = 480 'largura
		Jpeg.Height = 640 'altura
	Else
		tipo_img = "imagem na horizontal"
		Jpeg.Width = 640 'largura
		Jpeg.Height = 480 'altura
	End If
		Jpeg.Save salva & "\g_"&gera&".jpg" 'Imagem 800x600

	scale = 60
	Jpeg.Width = Jpeg.OriginalWidth * Scale / 100
	Jpeg.Height = Jpeg.OriginalHeight * Scale / 100
	Jpeg.Save salva & "\m_"&gera&".jpg" 'Imagem Média

	Scale = 35
	Jpeg.Width = Jpeg.OriginalWidth * Scale / 100
	Jpeg.Height = Jpeg.OriginalHeight * Scale / 100
	Jpeg.Save salva & "\p_"&gera&".jpg" 'Imagem Pequena
	
'converte em 800x600	

	File.Delete 'Deata a imagem
	'Set Fil=Server.CreateObject("Scripting.FileSystemObject") 
		 
	'	Set filObject=Fil.GetFile(fotos_dir & "\g_" & gera&".jpg") 
	'		 	filObject.Delete True
	temp_img = gera
	
%>
<font size="2" face="Verdana, Arial, Helvetica, sans-serif"> <strong>Seu croqui 
foi enviado com Sucesso!!! </strong><br>
Abaixo aparece uma amostra de como ficou o seu croqui. <br>
Caso<font color="#FF0000"><strong> nao tenha ficado bom</strong></font>, por favor 
<a href="?img=deleta&temp_img=<%=temp_img%>">clique para</a> enviar o seu croqui 
novamente, caso ainda tenha dúvidas por favor visite nossa sessao de dicas de 
como criar seu croqui! <br>
</font><font color="#006600" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
<input type="button" name="Submit" value="Conclu&iacute;do/Continuar" 
		onClick="javascript:location.href='fotos_editar_admin.asp?cadastrar=novo&upload=ok&temp_img=<%=temp_img%>'">
</strong></font> 
<hr>
<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><BR>
<BR>
Foto no tamanho que vai ampliar:<br>
<IMG SRC="../upload/fotos/m_<% = gera&".jpg"%>"><br>
Seu croqui em tamanho miniatura:<br>
<IMG SRC="../upload/fotos/p_<% = gera&".jpg"  %>"><br>
<%=request.QueryString("id")%>-<%=request.QueryString("ac")%>
<script>
location.href='fotos_editar_admin.asp?ac=<%=request.QueryString("ac")%>&id=<%=request.QueryString("id")%>&temp_img=<%=gera%>';
</script>
<%

end if
Next
eND IF
%>
<p align="center"><!--#include file="rodape.asp"-->