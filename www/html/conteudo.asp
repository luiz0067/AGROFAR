			<div class="linha">
			<%if request("pg")="produtos" then%>				
				<!--#include file="produtos.asp"-->		
						
			<%else %>
				<div id="conteudo" class="centro">
					<div style="width:464px;float:left;padding-left:16px;">					
					<!--------------coteudo--------------------------------------------->					
						<%if request("pg")="noticias" then%>				
							<!--#include file="noticias.asp"-->		
						<%elseif request("pg")="o_sistema_gps" then%>
							<!--#include file="o_sistema_gps.asp"-->		
						<%elseif request("pg")="sustentabilidade" then%>
							<!--#include file="sustentabilidade.asp"-->		
						<%elseif request("pg")="gps_e_agricultura" then%>
							<!--#include file="gps_e_agricultura.asp"-->
						<%elseif request("pg")="contato" then%>
							<!--#include file="contato.asp"-->
						<%elseif request("pg")="suporte" then%>
							<!--#include file="suporte.asp"-->
						<%elseif request("pg")="evolution" then%>
							<!--#include file="evolution.asp"-->
						<%else%>						
									
						<%end if%>
						
					<!--------------coteudo--------------------------------------------->
					</div>
					<div style="width:200px;float:right;" >
						<!------------------menu----------------------------------------->
						<%if request("pg")="contato" then%>
						<!--#include file="lateral_direita_contato.asp"-->	
						<%elseif request("pg")="suporte" then%>
						<!--#include file="manual_pdf.asp"-->	
						<%elseif request("pg")="evolution" then%>
						<!--#include file="lateral_direita_evolution.asp"-->
						<%else%>						
						<!--#include file="menu.asp"-->
						<%end if%>
						<!------------------menu----------------------------------------->
					</div>
				</div>
			</div>
			<%end if%>
