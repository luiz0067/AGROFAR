<%
set Rs = server.createobject("ADODB.recordset")
Sqltext = "SELECT fotos.foto as foto FROM fotos left join fotos_categorias on(fotos.id_cat=fotos_categorias.id)"
Sqltext = Sqltext&" where (fotos_categorias.categoria like '"&request("nome")&"')"
Rs.Open SqlText, Conn

%>							
			<div class="linha"  style="min-height:590px;height:auto;">
				<div  class="centro" style="width:700px;min-height:590px;height:auto;">
					<div class="linha" style="min-height:299px;height:auto;width:637px;">
						<div class="linha" style="min-height:90px;height:auto;width:637px;">
							<div class="centro" style="min-height:90px;height:auto;width:637px;" >
								<div id="mapaa"><br>
								  <a class="mapadepercurso1">MAPA DE PERCURSO</a><a class="mapadepercurso1"></a><br />
								 Você está em: Home > Produtos e Soluções > <%=request("nome")%><br>
								</div>
								<div class="linha" style="text-align:left;">
									<a style="font-family:hemihead;font-size:20px;color:#45196F;"><%=request("nome")%></a><br><br>
								</div>
							</div>
						</div>					
						<div class="linha" style="min-height:219px;height:auto;width:637px;">
							<div class="centro" style="width:637px;min-height:90px;height:auto;">
								<div style="min-height:219px;height:auto;width:9px;float:left"></div>
								<div style="min-height:219px;height:auto;width:432px;float:left;background-color:#EFEFEF">
									<div  style="min-height:219px;height:auto;width:171px;float:left">
										<div style="min-height:219px;height:auto;width:5px;float:left" ></div>
										<div style="min-height:219px;height:auto;float:left" >
											<div class="linha" style="height:5px;"></div>
											<div class="linha" style="height:20px;width:166px;background-image:url('./images/canto_slide_show.jpg');"></div>
											<div id="gallery" class="linha" style="height:158px;width:166px;">
													<%	
														contador=0
														Do Until (rs.EOF)																
															foto = rs("foto")
															
														
															if contador>0 then
																ocultar="display:none"
																id=""
															elseif contador=0 then
																id="foto_grande_produtos"
																ocultar=""
															end if
													%>
													<a href="../upload/fotos/g_<%=foto%>.jpg" title="<%=request("nome")%>" style="border-size:0px;<%=ocultar%>" >
														<img src="../upload/fotos/M_<%=foto%>.jpg" height="158px" width="166px" border="0px" id="<%=id%>">
													</a>
													<%		rs.MoveNext	
															contador=contador+1
														Loop 
													%>
											</div>							
											<div class="linha" style="height:6px;width:166px;background-image:url('./images/barra_cinza.jpg');overflow:hidden;"></div>								
											<div class="linha" style="width:166px;">
												<%
													contador=0
													rs.movefirst
													Do Until (rs.EOF)																
														foto = rs("foto")												%>
														<div CLASS="mini_foto"><img width="24px" height="27px" src="../upload/fotos/p_<%=foto%>.jpg" onclick="amplia('../upload/fotos/m_<%=foto%>.jpg')"></div>
												<%		
														rs.MoveNext	
														contador=contador+1
													Loop 
													if (contador<5) then
														Do Until (contador>5)
															%>
															<div CLASS="mini_foto"></div>
															<%
															contador=contador+1
														Loop 
													end if
												%>
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
								<div style="min-height:219px;height:auto;width:7px;float:left"></div>
								<div style="min-height:219px;height:auto;width:175px;float:left">
									<div style="min-height:100px;height:auto;width:175px;clear:both;background-color:#451970;color:#FFFFFF;font-family:eurostile;text-aling:left;font-size:14px;">
										<br>
										Entre em contato com<br>
										nosso suporte.<br>
										<br>
										+55 45 3054.8888
									</div>
									<div style="min-height:14px;height:auto;width:175px;clear:both;display:block"></div>
									<a href="?pg=contato" style="text-decoration:none">
										<div style="min-height:105px;height:auto;width:175px;clear:both;background-color:#A3C030;color:#FFFFFF;font-family:eurostile;text-aling:right;font-size:14px;">
											<br>
											<br>
											Solicite um orçamento<br>
											Clique Aqui.
										</div>
									</a>
								</div>
								<div style="min-height:219px;height:auto;width:14px;float:left"></div>
							</div>
						</div>
					</div>
					<div class="linha" style="min-height:229px;height:auto;">
						<div class="linha" style="min-height:86px;height:auto;">
							<div class="linha" style="min-height:48px;height:auto;"></div>
							<div class="linha" style="min-height:38px;height:auto;">
								<div class="centro"  style="min-height:38px;height:auto;width:680px">
									<div id="botoes" style="min-height:38px;height:auto;width:680px;display:none">
										<div class="botao_esquerda" onclick="anterior();"></div>				
										<div class="botao_direita" onclick="proximo();"></div>					
									</div>
								</div>
							</div>
						</div>					
						<div class="botao_esquerda" onclick="anterior();" 	style="position:absolute;margin-top:50px;"></div>
						<div class="botao_direita" onclick="proximo();" 	style="position:absolute;margin-top:50px;margin-left:660px;"></div>
						<div class="linha" style="background-image:url('images/fundo2.jpg');min-height:143px;height:auto;">
							<div id="rolagem" class="centro" style="height:143px;">					
								<div  style="height:143px;width:3000px;">
								
									<%
										'Sqltext = "SELECT fotos.foto,fotos_categorias.categoria as nome as foto FROM fotos left join fotos_categorias on(fotos.id_cat=fotos_categorias.id)"
										'Sqltext = Sqltext&" where (fotos_categorias.categoria like '"&request("nome")&"')"

										set Rs1 = server.createobject("ADODB.recordset")
										Sqltext = "SELECT fotos_categorias.id,fotos_categorias.categoria as nome FROM fotos_categorias"
										Rs1.Open SqlText, Conn
										Do Until rs1.EOF
											set Rs2 = server.createobject("ADODB.recordset")
											Sqltext = "SELECT fotos.foto FROM fotos where (id_cat="&rs1("id")&") order by id asc"
											Rs2.Open SqlText, Conn
											if not((rs2.EOF)and(rs2.BOF)) then
												%>
												<a href="?pg=produtos&nome=<%=rs1("NOME")%>" style="text-decoration:none;">
													<div class="foto_miniatura">
														<div style="font-family:hemihead;font-size:30px;color:#45196F;position:absolute;margin-top:90px;color:#A3C030;width:134px;text-align:center;height:32px;overflow:hidden;display:none"><%=rs1("NOME")%></div>
														<img src="../upload/fotos/M_<%=rs2("FOTO")%>.jpg" height="124px" width="134px">
													</div>												
												</a>
												<%
												rs2.Close
												Set rs2 = Nothing
											end if
											rs1.MoveNext
										Loop 
										rs1.Close
										Set rs1 = Nothing
									%>
								</div>
							</div>
						</div>
					</div>				
					<div class="linha" style="min-height:62px;height:auto;">
						<div class="centro" style="min-height:62px;height:auto;"></div>
					</div>
				</div>
			</div>				
			<!--andreza.torre@fornecedores.vivo.com.br-->
<%
rs.Close
Set rs = Nothing
conn.close
%>