<!--#include file="config_site.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title>
		
		<%if request("pg")="noticias" then%>				
			A Agricultura de Precisão
		<%elseif request("pg")="o_sistema_gps" then%>
			O sistema GPS
		<%elseif request("pg")="sustentabilidade" then%>
			O GPS e Agricultura de Precisão		
		<%elseif request("pg")="gps_e_agricultura" then%>
			GPS e agricultura
		<%elseif request("pg")="contato" then%>
			Contato
		<%elseif request("pg")="suporte" then%>
			Suporte
		<%elseif request("pg")="evolution" then%>
			Evolution
		<%elseif request("pg")="produtos" then%>
			Produtos <%=request("nome")%>				
		<%else%>						
					
		<%end if%>
		
		
		</title>
		<META http-equiv=Content-Type content="text/html">
		<link rel="stylesheet" type="text/css" href="css/jquery.lightbox-0.5.css" media="screen" />
		<style type="text/css">
			@import url("./css/estilos.css");
        </style>
			<link rel="stylesheet" type="text/css" href="../style-projects-jquery.css" />    
			<link rel="stylesheet" type="text/css" href="css/jquery.lightbox-0.5.css" media="screen" />
			<style type="text/css">
			/* jQuery lightBox plugin - Gallery style */
			#gallery {
			}
			#gallery ul { list-style: none; }
			#gallery ul li { display: inline; }
			#gallery ul img {
			}
			#gallery ul a:hover img {
			}
			#gallery ul a:hover { color: #fff; }
			</style>
			<script type="text/javascript" src="js/jquery.js"></script>
			<script type="text/javascript" src="js/jquery.lightbox-0.5.js"></script>
			<script type="text/javascript">
			$(function() {
				$('#gallery a').lightBox();
			});
			function buscar(id){
				return document.getElementById(id);
			}
			function load(){
				if (window.innerHeight)
					buscar("pagina").style.height=(f_clientHeight()-203)+"px";
				else
					buscar("pagina").style.height=(f_clientHeight()-218)+"px";
			}
			
			function f_clientHeight() {
				return ((window.innerHeight)? (window.innerHeight) : document.body.clientHeight);
			}
			function proximo(){
				roller=buscar("rolagem");
				limite=roller.childNodes.length*149;
				if (limite>roller.scrollLeft)
					roller.scrollLeft=roller.scrollLeft+143;
				else
					roller.scrollLeft=0;
			}
			function anterior(){
				roller=buscar("rolagem");
				if (0<roller.scrollLeft)
					roller.scrollLeft=roller.scrollLeft-143;
				else
					roller.scrollLeft=limite
			}
			 
			$(function() {
				$('#gallery a').lightBox();
			});
			function amplia(url){
				document.getElementById('foto_grande_produtos').src=url;
			}
		</script>
	</head>
	<body bgcolor="#ffffff" onload="load()">
		<div id="tudo" >
<!--#include file="topo.asp"-->
			<div id="pagina" class="linha" style="height:50px;">
				<!--#include file="conteudo.asp"-->
			</div>
<!--#include file="rodape.asp"-->			
		</div>
    </body>
</html>
