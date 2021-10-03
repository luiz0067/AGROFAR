//Prototype e herança de objeto
//crio um objeto principal e outros objetos que serao as subClasse para o primeiro, fazendo a herança de objetos
function ajax() {
};
//metodo iniciar
ajax.prototype.iniciar = function() {
//logo abaixo tentamos estanciar o objeto XMLHttpRequest() para o IE e os demais navegadores
    try{
        this.xmlhttp = new XMLHttpRequest();
    }catch(ee){
        try{
            this.xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
        }catch(e){
            try{
                this.xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
            }catch(E){
                this.xmlhttp = false;
            }
        }
    }
    return true;
}
//metodo oculpado
ajax.prototype.ocupado = function() {
    estadoAtual = this.xmlhttp.readyState;
    return (estadoAtual && (estadoAtual < 4));
}
//metodo processa
ajax.prototype.processa = function() {
    if (this.xmlhttp.readyState == 4 && this.xmlhttp.status == 200) {
        return true;
    }
}
//metodo envia esperando a url, o metodo get ou post e o modo true ou false
ajax.prototype.enviar = function(url, metodo, modo) {
//se nao existir o objeto xmlhttp ele sera criado com o metodo iniciar
    if (!this.xmlhttp) {
        this.iniciar();
    }
//se nao estiver oculpado
    if (!this.ocupado()) {
//se o metodo passado for get
        if(metodo == "GET") {
            this.xmlhttp.open("GET", url, modo);
            this.xmlhttp.send(null);
        } else {
//se for post        
            this.xmlhttp.open("POST", url, modo);
            this.xmlhttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded; charset=UTF-8");
            this.xmlhttp.setRequestHeader("Cache-Control", "no-store, no-cache, must-revalidate");
            this.xmlhttp.setRequestHeader("Cache-Control", "post-check=0, pre-check=0");
            this.xmlhttp.setRequestHeader("Pragma", "no-cache");
            this.xmlhttp.send(url);
        }    
//se estiver ok ele retorna o resultado e ja utilizamos um modo para receber acentos
// usando a função unescape e substituindo os + por espaços mesma coisa que usar urldecode() do php
        if (this.processa) {
//recebe o resultado da da pagina php
            return unescape(this.xmlhttp.responseText.replace(/\+/g," "));
        }
    }
    return false;
}