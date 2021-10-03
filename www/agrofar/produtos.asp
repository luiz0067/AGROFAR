
			<div class="linha"  >
				
				<div class="linha"style="height:90px;">
					<div class="centro" style="width:637px;" >
						<div id="mapaa"><br>
						  <a class="mapadepercurso1">MAPA DE PERCURSO</a><a class="mapadepercurso1"></a><br />
						 Você está em: Home > Produtos e Soluções > <%=request("nome")%><br>
						</div>
						<div class="linha" style="text-align:left;">
							<a style="font-family:hemihead;font-size:20px;color:#45196F;"><%=request("nome")%></a><br><br>
						</div>
                	</div>
				</div>
				<div class="linha" style="height:219px;">
					<div class="centro" style="width:637px">
						<div style="height:219px;width:9px;float:left"></div>
						<div style="min-height:219px;height:auto;width:432px;float:left;background-color:#EFEFEF">
							<div  style="height:219px;width:171px;float:left">
								<div style="height:219px;width:5px;float:left" ></div>
								<div style="height:219px;float:left" >
									<div class="linha" style="height:5px;"></div>
									<div class="linha" style="height:20px;width:166px;background-image:url('./images/canto_slide_show.jpg');"></div>
									<div id="gallery" class="linha" style="height:158px;width:166px;">
										<%if request("nome")="AP 100" then%>	
											<a href="./photos/ap100_g1.jpg" title="<%=request("nome")%>" style="border-size:0px;">
												<img src="./images/AP100.jpg" height="158px" width="166px" border="0px">
											</a>
										<%elseif request("nome")="AP 200" then%>
											<a href="./photos/ap200_g1.jpg" title="<%=request("nome")%>" style="border-size:0px;">
												<img src="./images/AP200.jpg" height="158px" width="166px" border="0px">
											</a>
										<%elseif request("nome")="AP 300" then%>
											<a href="./photos/ap300_g1.jpg" title="<%=request("nome")%>" style="border-size:0px;">
												<img src="./images/AP300.jpg" height="158px" width="166px" border="0px">
											</a>
										<%elseif request("nome")="AP 400" then%>
											<a href="./photos/ap400_g1" title="<%=request("nome")%>" style="border-size:0px;">
												<img src="./images/AP400.jpg" height="158px" width="166px" border="0px">
											</a>
										<%elseif request("nome")="GPS GUIA AGROFAR" then%>
											<a href="./photos/gpsguia_g1.jpg" title="<%=request("nome")%>" style="border-size:0px;">
												<img src="./images/GPSGuia.jpg" height="158px" width="166px" border="0px">
											</a>

										<%end if%>
									</div>							
									<div class="linha" style="height:6px;width:166px;background-image:url('./images/barra_cinza.jpg');overflow:hidden;"></div>								
									<div class="linha" style="width:166px;">
										<div CLASS="mini_foto"></div>									
										<div CLASS="mini_foto"></div>									
										<div CLASS="mini_foto"></div>									
										<div CLASS="mini_foto"></div>									
									</div>									
								</div>						
							</div>
							<div style="min-height:219px;height:auto;width:261px;float:left;">
								<div class="linha" style="height:18px"></div>
								<div class="linha" >
									<div style="width:18px;height:100%;float:left;"></div>
									<div style="float:left;text-align:justify;" id="coluna_produtos" >
										<%if request("nome")="AP 100" then%>				
											<!--#include file="AP100.asp"-->
										<%elseif request("nome")="AP 200" then%>
											<!--#include file="AP200.asp"-->
										<%elseif request("nome")="AP 300" then%>
											<!--#include file="AP300.asp"-->	
										<%elseif request("nome")="AP 400" then%>
											<!--#include file="AP400.asp"-->	
										<%elseif request("nome")="GPS GUIA AGROFAR" then%>
											<!--#include file="GPSGUIA.asp"-->	
										<%end if%>								
									</div>
								</div>
							</div>
						</div>
						<div style="height:219px;width:7px;float:left"></div>
						<div style="height:219px;width:175px;float:left">
							<div style="height:100px;width:175px;clear:both;background-color:#451970;color:#FFFFFF;font-family:eurostile;text-aling:left;font-size:14px;">
								<br>
								Entre em contato com<br>
								nosso suporte.<br>
								<br>
								+55 45 3054.8888
							</div>
							<div style="height:14px;width:175px;clear:both;display:block">
							</div>
							<a href="?pg=contato" style="text-decoration:none">
								<div style="height:105px;width:175px;clear:both;background-color:#A3C030;color:#FFFFFF;font-family:eurostile;text-aling:right;font-size:14px;">
									<br>
									<br>
									Solicite um orçamento<br>
									Clique Aqui.
								</div>
							</a>
						</div>
						<div style="height:219px;width:14px;float:left"></div>
					</div>
				</div>
				<div class="linha" style="height:48px;"></div>
				<div class="linha" style="height:38px;">
					<div class="centro" style="width:680px">
						<div id="botoes" style="width:680px">
							<div class="botao_esquerda" onclick="anterior();"></div>				
							<div class="botao_direita" onclick="proximo();"></div>					
						</div>
					</div>
				</div>
				<div class="linha" style="background-image:url('images/fundo2.jpg');">
					<div id="rolagem" class="centro">					
						<div  style="height:143px;width:3000px;">
								<a href="?pg=produtos&nome=AP 100"><div class="foto_miniatura"><img src="images/mini_AP100.jpg" height="124px" width="134px"></div></a>
								<a href="?pg=produtos&nome=AP 200"><div class="foto_miniatura"><img src="images/mini_AP200.jpg" height="124px" width="134px"></div></a>
								<a href="?pg=produtos&nome=AP 300"><div class="foto_miniatura"><img src="images/mini_AP300.jpg" height="124px" width="134px"></div></a>
								<a href="?pg=produtos&nome=AP 400"><div class="foto_miniatura"><img src="images/mini_AP400.jpg" height="124px" width="134px"></div></a>
								<a href="?pg=produtos&nome=GPS GUIA AGROFAR"><div class="foto_miniatura"><img src="images/mini_GPS_Guia.jpg" height="124px" width="134px"></div></a>
						</div>
					</div>
				</div>
				<div class="linha">
					<div class="centro" style="width:62px"></div>
				</div>
			</div>

			