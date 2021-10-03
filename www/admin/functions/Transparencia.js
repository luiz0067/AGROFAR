	function aparecer(inicio,fim,tempo,grau){
		transparencia('hint',inicio)
		inicio=inicio+grau
		if (inicio>fim+2){
			return 0
		}
		else{
			setTimeout('aparecer('+inicio+','+fim+','+tempo+','+grau+')',tempo);
		}
	}

	function ocultar(inicio,fim,tempo,grau){
		transparencia('hint',fim)
		fim=fim-grau
		if (inicio-30>fim){
			return 0
		}
		else{
			setTimeout('ocultar('+inicio+','+fim+','+tempo+','+grau+')',tempo);
		}
	}
 	function transparencia(objeto,valor){
		document.getElementById(objeto).style.filter="alpha(opacity="+valor+")"
		if (valor=0)
			document.getElementById(objeto).style.display='block'
	}
	function MostrarDetalhes(link_da_pagina){
		document.getElementById("hint").style.display='block'
		aparecer(1,100,50,30)
		//document.getElementById("pagina").location.href=link_da_pagina;
		pagina.location.href=link_da_pagina;
	}
	function OcultarDetalhes(){
			ocultar(1,100,50,30)
	}