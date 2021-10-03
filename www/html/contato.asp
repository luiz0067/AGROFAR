		
		<div id="coluna1" >	
		
		
		
			<div class="linha"style="height:60px;">
				<div id="mapaa"><br>
				  <a class="mapadepercurso1">MAPA DE PERCURSO</a><a class="mapadepercurso1"></a><br />
				  Você está em: Home &gt; Entre em contato<br>
				</div>
			</div>		
			<div class="linha" style="text-align:left">
				<a style="font-family:hemihead;font-size:20px;color:#45196F;text-align:justify;">Entre em contato</a><br><br>
				<br>
			</div>
			
			<div class="arrendondado" style="width:100%;background-color:#FAFAFA;">

				<div  class="linha">
					<div  class="centro" style="width:320px">
						<div  class="linha" id="formulario" style="text-align:left;">				
							<form  method="post" name="formulario"  style="font-family:eurostile;font-size:12px" >
							<br />
							Nome:<br />
							<input name="nome" type="text" size="50" style="width:320px" />
							<br />
							E-mail:<br />
							<input name="email" type="text" size="50" style="width:320px" />
							<br />
							<br />
							Telefone:<br />
							<input name="telefone" type="text" size="50"  style="width:320px"/>
							<br />
							<br />
							Cidade:<br />
							<input name="cidade" type="text" size="50" style="width:320px" />
							<br />
							<br />
							Mensagem:<br />
							<textarea name="mensagem:" cols="50" rows="8" style="width:320px"></textarea>
							<br />
							<br />
							<input name="acao" type="submit" value="ENVIAR" />
							<input name="acao"  type="submit" value="LIMPAR" />
							<br />
							<br />
							</form>
						</div>
					</div>
				</div>
			</div>
	<%
	'set AspEmail=""
	'set nomeRemetente=""
	'set emailRemetente=""
	'set nomeDestinatario=""
	'set emailDestinatario=""
	'set emailRetorno=""
	'set assunto=""
	'set mensagem=""
	'set servidor=""
	if(request("acao")="ENVIAR") then
	
		'Declaramos as váriaveis a serem utilizadas no script
		 
		'Configuramos os dados a serem utilizados no cabeçalho da mensagem
		nomeDestinatario="suporte"
		emailDestinatario="suporte@agrofar.com.br"
		nomeRemetente=request("nome")
		emailRemetente="suporte@agrofar.com.br"
		emailRetorno="suporte@agrofar.com.br"
		responderPara="suporte@agrofar.com.br"
		assunto="Contato de "&request("nome")
		
		mensagem = ""
		mensagem = mensagem&"<br>Nome............. " & request("nome") & vbcrlf
		mensagem = mensagem&"<br>E-mail........... " & request("email") & vbcrlf
		mensagem = mensagem&"<br>Telefone......... " & request("telefone") & vbcrlf
		mensagem = mensagem&"<br>Cidade........... " & request("cidade") & vbcrlf
		mensagem = mensagem&"<br>Mensagem......... " & request("mensagem") & vbcrlf

		servidor="localhost"
		 
		'Agora configuramos o componente utilizando os dados informados nas variáveis
		 
		'Instancia o objeto na memória
		SET AspEmail = Server.CreateObject("Persits.MailSender")
		 
		'Contfigura o servidor SMTP a ser utilizado
		AspEmail.Host = servidor
		 
		'Configura o Nome do remetente da mensagem
		AspEmail.FromName = nomeRemetente
		 
		'Configura o e-mail do remetente da mensagem que OBRIGATORIAMENTE deve ser um e-mail do seu próprio domínio
		AspEmail.From = emailRemetente
		 
		'Configura o E-mail de retorno para você ser avisado em caso de problemas no envio da mensagem
		AspEmail.MailFrom = emailRemetente
		 
		 
		'Configura o e-mail que receberá as respostas desta mensagem
		AspEmail.AddReplyTo responderPara
		 
		'Configura os destinatários da mensagem
		AspEmail.AddAddress emailDestinatario, nomeDestinatario
		 
		'Configura para enviar e-mail Com Cópia
		'AspEmail.AddCC "nome0@dominio.com.br", "Nome"
		'AspEmail.AddCC "nome1@dominio.com.br", "Nome"
		'AspEmail.AddCC "nome2@dominio.com.br", "Nome"
		 
		'Configura o Assunto da mensagem enviada
		AspEmail.Subject = assunto
		 
		'Configura o formato da mensagem para HTML
		AspEmail.IsHTML = True
		 
		'Configura o conteúdo da Mensagem
		AspEmail.Body = mensagem
		 
		'Utilize este código caso queira enviar arquivo anexo
		'AspEmail.AddAttachment("E:\home\SEU_LOGIN_FTP\Web\caminho_do_arquivo")
		 
		'Para quem utiliza serviços da REVENDA conosco
		'AspEmail.AddAttachment("E:\vhosts\DOMINIO_COMPLETO\httpdocs\caminho_do_arquivo")
		 
		'#Ativa o tratamento de erros
		On Error Resume Next
		 
		'Envia a mensagem
		AspEmail.Send
		 
		'Caso ocorra problemas no envio, descreve os detalhes do mesmo.
		If Err <> 0 Then
			erro = "<b><font color='red'> Erro ao enviar a mensagem.</font></b><br>"
			erro = erro & "<b>Erro.Description:</b> " & Err.Description & "<br>"
			erro = erro & "<b>Erro.Number:</b> "      & Err.Number & "<br>"
			erro = erro & "<b>Erro.Source:</b> "      & Err.Source & "<br>"
			response.write erro
		Else
			%>
			<a class="titulo" style="font-family:hemihead;font-size:20px;text-align:justify;color:#45196F;">
				Sua mensagem enviada com sucesso!
			</a><br><br>
			<%
		End If
		 
		'## Remove a referência do componente da memória ##
		SET AspEmail = Nothing
	
	End If
%>


		</div>