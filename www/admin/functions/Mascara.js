
//TextRange

function selectSomeText(element,pos)
{
	if (element.setSelectionRange)
	{
		element.setSelectionRange(begin,end);
	}
	else if (element.createTextRange)
	{
		var range = element.createTextRange();
		//range.moveStart("character",begin);
		//range.moveEnd("character",end);
		range.move("character", pos); 
		range.select();
	}
}

	var Mask
	function setMask(m){
		Mask=m
	}
	function getMask(){
		//Mask="000"
		return Mask
	}

	function mascara(mascara,campo){

		//mascara.length
		var texto;
		var nulo;
		var completa;
		var ver;
		var regra;
		 texto=campo.value	
			 texto=texto.replaceAll("_","")	
		 nulo=mascara.replaceAll("0","_") 		
		 
		x=0  
		while(true){			
			if (mascara.substring(x,x+1)=="0"){
				 regra= /^[0-9]{1}$/i;
			}
			if (mascara.substring(x,x+1)=="."){
				 regra= /^[.]{1}$/i;
			}
			else if (mascara.substring(x,x+1)=="a"){	
				regra=/^[a-z]{1}$/i;
			}
			else if (mascara.substring(x,x+1)=="-"){	
				regra=/^[-]{1}$/i;
			}
			else if (mascara.substring(x,x+1)=="/"){	
				regra=/^[./-]{1}$/i;
			}
			ver = regra.exec(texto.substring(x,x+1));			
			
			if(!ver){
				texto=texto.substring(0,x)+"_".substring(x+1,texto.length)
			}
			x++	
			
			if ((x>texto.length)&&(x>mascara.length))
				break;
		}

		texto=texto.substring(0,mascara.length)
		texto=texto.trim();
		completa=nulo.substring(texto.length,mascara.length)
		campo.value=texto+completa
		selectSomeText(campo,texto.length)	
		transfere()
	}




	function TornarCPF(){
		document.getElementById('CPF_CGC').innerHTML=' <strong><font color=#FFFFFF size=1 face=Verdana, Arial, Helvetica, sans-serif>::   Buscar clientes por CPF::</font> </strong>'	;
		document.getElementById('maskedit').disabled = 0; // para ativado
//		document.getElementById('maskedit').value="12346542145";
		document.getElementById('maskedit').value="___.___.___-__"
//		document.getElementById('maskedit').setEvent("KeyUp",mascara('00.000.000/0000-00',maskedit))
		setMask("000.000.000-00")
    
	}
	function TornarCGC(){
		document.getElementById('CPF_CGC').innerHTML=' <strong><font color=#FFFFFF size=1 face=Verdana, Arial, Helvetica, sans-serif>::   Buscar clientes por CGC::</font> </strong>'
		document.getElementById('maskedit').disabled = 0; // para ativado
		document.getElementById('maskedit').value="__.___.___/____-__"
		
		//000.000.000-00
		//document.form1.nome.disabled = 1; // para desativado		
		setMask("00.000.000/0000-00")
	}


function transfere(){
	document.getElementById('cpf_cgc_dado').value=document.getElementById('maskedit').value.replaceAll('.','').replaceAll('-','').replaceAll('/','').replaceAll('_','')
	
}