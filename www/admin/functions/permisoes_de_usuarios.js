function cad(id){
pg = id;
	window.open(pg,'WinExa','toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=0,width=210,height=5')
}

          function mascara_niver(niver){ 
              var mydata = ''; 
              mydata = mydata + niver; 
              if (mydata.length == 2){ 
                  mydata = mydata + '/'; 
                  document.config.niver.value = mydata; 
              } 
              if (mydata.length == 5){ 
                  mydata = mydata + '/'; 
                  document.config.niver.value = mydata; 
              } 
              if (mydata.length == 10){ 
                 verifica_data(); 
				 ano = (document.config.niver.value.substring(6,10))
				 ano = new Date(new Date()).getFullYear() - ano;
				 document.config.idade.value = ano;
				 
              } 
          } 
           
          function verifica_data () { 

            dia = (document.config.niver.value.substring(0,2)); 
            mes = (document.config.niver.value.substring(3,5)); 
            ano = (document.config.niver.value.substring(6,10)); 

            situacao = ""; 
            // verifica o dia valido para cada mes 
            if ((dia < 01)||(dia < 01 || dia > 30) && (  mes == 04 || mes == 06 || mes == 09 || mes == 11 ) || dia > 31) { 
                situacao = "falsa"; 
            } 

            // verifica se o mes e valido 
            if (mes < 01 || mes > 12 ) { 
                situacao = "falsa"; 
            } 

            // verifica se e ano bissexto 
            if (mes == 2 && ( dia < 01 || dia > 29 || ( dia > 28 && (parseInt(ano / 4) != ano / 4)))) { 
                situacao = "falsa"; 
            } 
    
            if (document.config.niver.value == "") { 
                situacao = "falsa"; 
            } 
    
            if (situacao == "falsa") { 
                alert("Data inválida!"); 
                document.config.niver.focus(); 
            } 
          } 

//pega o ano corrente
//document.write(escreve_ano());


//************************************
// Mascaras para CEP e Telefone

//************************************
          function mascara_cep(cep){ 
              var mycep = ''; 
              mycep = mycep + cep; 
              if (mycep.length == 5){ //conta Cinco digitos
                  mycep = mycep + '-'; 
                  document.config.cep.value = mycep; 
              } 
              if (mycep.length == 9){ 
                  //alert("CEP foi digitado errado,\n o CEP deve ser nnnnn-nnn.\n n = número de 0 à 9.");
				  //document.config.cep.focus();
				  frames.city2.cidade.focus(); 
             } 
          } 
		  function mascara_fone(fone){ 
              var myfone = ''; 
              myfone = myfone + fone; 
              if (myfone.length == 1){ //conta inicio do fone
                  myfone = '(' + myfone; //coloca entrada do DDD
                  document.config.fone.value = myfone; 
              }
			   if (myfone.length == 3){ //conta inicio do fone
                  myfone = myfone + ') '; //coloca final  do DDD mais um espaço
                  document.config.fone.value = myfone; 
              } 
			   if (myfone.length == 9){ //conta (00) 0000
                  myfone = myfone + '-'; //coloca entrada do DDD
                  document.config.fone.value = myfone; 
              } 
                  if (myfone.length == 14){ // Avisa se digitou mais que um telefone padrao 
                  //alert("Telefone foi digitado errado,\n o Telefone deve ser (dd) nnnn-nnnn.\n n = número de 0 à 9.");
				 
				  document.config.fone_cel.focus(); 
             } 
          }
		    function mascara_fone_cel(fone_cel){ 
              var myfone_cel = ''; 
              myfone_cel = myfone_cel + fone_cel; 
              if (myfone_cel.length == 1){ //conta inicio do fone
                  myfone_cel = '(' + myfone_cel; //coloca entrada do DDD
                  document.config.fone_cel.value = myfone_cel; 
              }
			   if (myfone_cel.length == 3){ //conta inicio do fone
                  myfone_cel = myfone_cel + ') '; //coloca final  do DDD mais um espaço
                  document.config.fone_cel.value = myfone_cel; 
              } 
			   if (myfone_cel.length == 9){ //conta (00) 0000
                  myfone_cel = myfone_cel + '-'; //coloca entrada do DDD
                  document.config.fone_cel.value = myfone_cel; 
              } 
                  if (myfone_cel.length == 14){ // Avisa se digitou mais que um telefone padrao 
                  //alert("Telefone foi digitado errado,\n o Telefone deve ser (dd) nnnn-nnnn.\n n = número de 0 à 9.");
				 
				  document.config.padrinho.focus(); 
             } 
          } 
		  function limpa_fone(){
			if (document.config.fone.value == "(dd) nnnn-nnnn")
				document.config.fone.value = "";
			}

function atualiza_campos(){
		document.config.profissao.value = frames.city.profissao2.value; 
		document.config.cidade.value = frames.city2.cidade.value; 
		document.config.veiculo_atual.value = frames.veiculo_atual.veiculos.value; 
		document.config.veiculo_prete.value = frames.veiculo_prete.veiculos_p.value;
		}
