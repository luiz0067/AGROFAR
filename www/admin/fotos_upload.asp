<!--#include file="login_nivel2.asp"-->
<% 

logo_dir = logo_dir 'Server.MapPath(".")
tamanho = 307200 
pode = tamanho / 1024 

if request("img") = "deleta" Then
on error resume next
	Set Fil=Server.CreateObject("Scripting.FileSystemObject") 
	Set filObject=Fil.GetFile(logo_dir & "/m_" & request("temp_img")&".jpg") 
		filObject.Delete True
	Set filObject=Fil.GetFile(logo_dir & "/p_" & request("temp_img")&".jpg") 
		filObject.Delete True
	Set filObject=Fil.GetFile(logo_dir & "/g_" & request("temp_img")&".jpg") 
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
							<div align="center" class="f_vermelho_3_b">Selecione a Foto/Imagem para enviar:</div></td>
						</tr>
						<tr>
							<td align="center" bgcolor="#F0F0F0"> 
								<font size="1" face="Arial, Helvetica, sans-serif">
									<img src="../imagens/icos_e_fundos/info.png" width="16" height="16" align="absmiddle"> 
									O Sistema s&oacute; aceita imagem enviados em JPG.
								</font><br> 
								<font size="1" face="Arial, Helvetica, sans-serif">
									<img src="../imagens/icones/picture.png" width="16" height="16" align="absmiddle"> 
									O tamanho m&aacute;ximo do arquivo enviado &eacute; de <%=Formatnumber(fix(tamanho/1024),0)%> kb.
								</font>
								<br> 
								<input type="FILE" size="30" name="FILE1">
								<br>
								<input name="SUBMIT" type=SUBMIT class="f_vermelho_3" style="color: #ff0000; font-family: Arial;" value="Enviar Foto!">
								<br>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</FORM>
<%
End If
if request("up") = "send" Then
	Set Upload = Server.CreateObject("Dundas.Upload.2")
	Upload.MaxFileSize = 3000000
	Upload.UseUniqueNames = False
	Set Jpeg = Server.CreateObject("Persits.Jpeg") 
	If IsObject(upload) Then
		FExtension = MID(FName,instr(FName,".")+1)
		upload.Save logo_dir&"/"
		FName = Mid(Upload.Files(0).Path, InstrRev(Upload.Files(0).Path, "/") + 1)
		FCONT = upload.Files(0).ContentType
		If upload.Files(0).Size > tamanho then
		%>
			<table width="650" border="0" cellspacing="1" cellpadding="1" align="center" class="moldura1">
				<tr>
					<td height="21">
						<p align="center">
							<img src="../imagens/icos_e_fundos/warning.png" align="absmiddle"> 
							<span class="f_vermelho_3_b" style="color: #FF0000; font-weight: bold">ERRO AO ENVIAR ARQUIVO!</span>
						</p>
						<p>
							<span class="f_vermelho_3_b">O arquivo � muito Pesado.</span><br>
							Desculpe mas o arquivo que voce est� tentando enviar � muito pesado, o 
							<strong>m�ximo</strong> permitido � de<strong> 300 kb</strong>.<br>
							O arquivo que voce est� enviando tem (<strong><%=Formatnumber(fix(file.size/1024),0)%> kb</strong>).<br />
							Por favor tente novamente!!!<br>
							<br>
							<img src="../imagens/icos_e_fundos/question.png" alt="Ajuda" width="16" height="16" align="absmiddle"> 
							Esta com <span style="font-weight: bold">problemas</span> para reduzir o 
							<span style="font-weight: bold">tamanho</span> da sua <span style="font-weight: bold">foto</span> , ou entao quer 
							<span style="font-weight: bold">converte-la</span> em formato 
							<span style="font-weight: bold">JPG</span>?<br> 
							<img src="../imagens/icos_e_fundos/question.png" alt="Ajuda" width="16" height="16" align="absmiddle"> Visite o site 
							<a href="http://media-convert.com" target="_blank" class="f_vermelho_1_b">http://media-convert.com</a> e converta seus arquivos online de maneira r�pida e simples. <br>
							<img src="../imagens/icos_e_fundos/question.png" alt="Ajuda" width="16" height="16" align="absmiddle"> O tamanho que usamos � 
							de <span style="font-weight: bold">largura</span> (Width) = 
							<span style="font-weight: bold">480</span>, <span style="font-weight: bold">altura</span> (Height) = 
							<span style="font-weight:bold">640</span> para imagens na <span style="font-weight: bold">vertical</span>
							, para <span style="font-weight:bold">horizontal</span> <span style="font-weight: bold">largura</span> (Width) = 
							<span style="font-weight: bold">640</span>,altura (Height) = <span style="font-weight: bold">480</span><br /><br>						
							<a href="javascript:history.go(-1);"><img src="../imagens/icos_e_fundos/previous.png" alt="Voltar" border="0"  align="absmiddle"/>VOLTAR</a></font>
							<br>
							<br>
						</p>
					</td>
				</tr>
			</table>
		<%
		elseIf (FCONT="image/jpeg")or(FExtension = "jpg") OR (FExtension = "jpeg") Then
			Set Jpeg = Server.CreateObject("Persits.Jpeg")
			randomize timer	
			gera = (9876989856 * CInt((RND * 32000) + 100 ))
			Jpeg.Open logo_dir&"/"&FName
			Interpolation =	1
			Jpeg.Canvas.Font.Color = &000000' red
			Jpeg.Canvas.Font.BkMode = "Transparent"
			Jpeg.Canvas.Font.Family = "ARIAL"
			Jpeg.Canvas.Font.Size = 12
			'jpeg.Canvas.Font.Opacity = 0.9
			'Jpeg.Canvas.Font.Bold = True
			'Jpeg.Canvas.Print 2, 2, "�"&year(now)&" "&site_titulo
			if Jpeg.OriginalHeight > Jpeg.OriginalWidth Then
				tipo_img = "Imagem na vertical"
				Jpeg.Width = 480 'largura
				Jpeg.Height = 640 'altura
				Jpeg.Save logo_dir & "/g_"&gera&".jpg" 'Imagem 640x480
				Jpeg.Width = 158 'largura
				Jpeg.Height = 166 'altura
				Jpeg.Save logo_dir & "/m_"&gera&".jpg" 'Imagem 166x158
				Jpeg.Width = 24 'largura
				Jpeg.Height = 27 'altura
				Jpeg.Save logo_dir & "/p_"&gera&".jpg" 'Imagem 24x27
			Else
				tipo_img = "imagem na horizontal"
				Jpeg.Width = 640 'largura
				Jpeg.Height = 480 'altura
				Jpeg.Save logo_dir & "/g_"&gera&".jpg" 'Imagem 640x480
				Jpeg.Width = 166 'largura
				Jpeg.Height = 158'altura
				Jpeg.Save logo_dir & "/m_"&gera&".jpg" 'Imagem 166x158
				Jpeg.Width = 27 'largura
				Jpeg.Height = 24 'altura
				Jpeg.Save logo_dir & "/p_"&gera&".jpg" 'Imagem 24x27				
			End If					
			jpeg.close
			Set fs=Server.CreateObject("Scripting.FileSystemObject")
			set fs=nothing
			temp_img = gera
				%>
				<div>
					<font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
						<strong>Sua imagem foi enviada com Sucesso!!! </strong><br>
						Abaixo aparece uma amostra de como ficou a sua image. <br>
						Caso
						<font color="#FF0000"><strong> nao tenha ficado bom</strong></font>
						, por favor 
						<a href="?img=deleta&temp_img=<%=temp_img%>">clique para</a> 
						enviar a sua image 
						novamente, caso ainda tenha d�vidas por favor visite nossa sessao de dicas de 
						como criar seu croqui! <br>
					</font>
					<font color="#006600" size="2" face="Verdana, Arial, Helvetica, sans-serif">
						<strong> 
							<a href="fotos_editar_admin.asp?cadastrar=novo&upload=ok&temp_img=<%=temp_img%>&id=<%=request.QueryString("id")%>">
								<input type="button" name="Submit" value="Conclu&iacute;do/Continuar" onClick="javascript:location.href='fotos_editar_admin.asp?cadastrar=novo&upload=ok&temp_img=<%=temp_img%>'&id=<%=request.QueryString("id")%>">
							</a>
						</strong>
					</font> 
					<hr>
						<font size="2" face="Verdana, Arial, Helvetica, sans-serif">
							<BR>
							<BR>
							Foto no tamanho que vai ampliar:<br>
							<IMG SRC="../upload/fotos/m_<% = gera&".jpg"%>"><br>
							Sua image tamanho miniatura:<br>
							<IMG SRC="../upload/fotos/p_<% = gera&".jpg"  %>"><br>
							<%=request.QueryString("id")%>-<%=request.QueryString("ac")%>
							<script>
								location.href='fotos_editar_admin.asp?ac=<%=request.QueryString("ac")%>&id=<%=request.QueryString("id")%>&temp_img=<%=gera%>';
							</script>
						</font>
					</hr>
				</div>
		<%
		Else
			Set fs=Server.CreateObject("Scripting.FileSystemObject")
			if(fs.FileExists(logo_dir&"/"&FName))then
				fs.DeleteFile logo_dir&"/"&FName , True
			end if
			%>							
			<table width="650" border="0" cellspacing="1" cellpadding="1" align="center" class="moldura1">
				<tr>
					<td height="21">
						<p align="center">
							<img src="../imagens/icos_e_fundos/warning.png" align="absmiddle"> 
							<span class="f_vermelho_3_b" style="color: #FF0000;font-weight: bold">ERRO AO ENVIAR ARQUIVO!</span>
						</p> 
						<p>
							<span class="f_vermelho_3_b">Formato/Extensao do arquivo nao permitido.</span><br>
							Desculpe mas o arquivo que voce est� enviando nao � 
							um arquivo JPEG (<strong>.jpg</strong>) � um arquivo (<strong><%=FExtension%></strong>).<br />
							Por questao de seguran�a, nosso sistema s� permite o envio de arquivos neste formato.<br/>
							<br>
							<br>
							<img src="../imagens/icos_e_fundos/question.png" alt="Ajuda" width="16" height="16" align="absmiddle"> Esta com <span style="font-weight: bold">problemas</span> para reduzir 
							o <span style="font-weight: bold">tamanho</span> da sua 
							<span style="font-weight: bold">foto</span>, ou entao quer 
							<span style="font-weight: bold">converte-la</span> em formato 
							<span style="font-weight: bold">JPG</span>?<br> 
							<img src="../imagens/icos_e_fundos/question.png" alt="Ajuda" width="16" height="16" align="absmiddle"> Visite o site 
							<a href="http://media-convert.com" target="_blank" class="f_vermelho_1_b">http://media-convert.com</a> 
							e converta seus arquivos online de maneira r�pida e simples. <br>
							<img src="../imagens/icos_e_fundos/question.png" alt="Ajuda" width="16" height="16" align="absmiddle"> 
							O tamanho que usamos � de <span style="font-weight:bold">largura</span> (Width) = 
							<span style="font-weight: bold">480</span>, <span style="font-weight:bold">altura</span> 
							(Height) = <span style="font-weight: bold">640</span> para imagens na <span style="font-weight: bold">vertical</span>, para 
							<span style="font-weight: bold">horizontal</span> 
							<span style="font-weight: bold">largura</span> (Width) = 
							<span style="font-weight: bold">640</span>, altura (Height) = <span style="font-weight: bold">480</span><br />
							<br>
							<a href="javascript:history.go(-1);"><img src="../imagens/icos_e_fundos/previous.png" alt="Voltar" border="0"  align="absmiddle"/>VOLTAR</a></font><br><br>
						</p>
					</td>
				</tr>
			</table>
			<%
			Response.End
		End If
		<!--#include file="rodape.asp"-->
		Response.End
	End If
	Set Up = Nothing 
eND IF
%>
<p align="center"><!--#include file="rodape.asp"-->