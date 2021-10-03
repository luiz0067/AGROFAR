<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title>
		
		fonte
		
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
		</script>
	</head>
	<body bgcolor="#ffffff" onload="load()">
		<div id="tudo" >
			<div style="clear:both;font-family: 'essai';">esai</div>
			<div style="clear:both;font-family: 'hemihead';">hemi</div>
			<div style="clear:both;font-family:'eurostile';">eurostile</div>
			<div style="clear:both;">normal</div>
		</div>
    </body>
</html>
