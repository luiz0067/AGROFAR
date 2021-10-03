function erro(){
	   	msg = "           SESSÃO RESTRITA À USUÁRIOS REGISTRADOS               \n";
		msg += "________________________________________________________________\n\n\n";
		msg += "Esta sessão é restrita para usuários registrado no site da Dominiun .\n\n";
		msg += "Caso você já esteja registrado no nosso site, basta usar o formulário\n";
		msg += "do seu lado direito IDENTIFICAR-SE entrando com seu login e senha.\n\n";
		msg += "Caso não seja registrado basta clicar no botão Cadastrar que fica no\n";
		msg += "final do formulário IDENTIFICAR-SE.\n";
		msg += "Cadastre-se no nosso site é gratuito e você terá inumeras vantagens.\n";
		
		
		alert(msg + "\n\n");
		history.back(-1);
		//location.href="default.asp";
}
function FormatText(command, option){
	
  	frames.message.document.execCommand(command, true, option);
  	frames.message.focus();
}
function Add(option){	
	
			frames.message.document.body.innerHTML =  frames.message.document.body.innerHTML + option; 
	
  frames.message.focus();
	}
function foreColor()	{
	var arr = showModalDialog("select_color.html",window,"font-family:Verdana; font-size:12; status:No; dialogWidth:22; dialogHeight:21" );
	if (arr != null) cmdExec("ForeColor",arr);	

}
function upload()	{
	window.open('news_upload.asp','WinExa','toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=0,width=400,height=150')
}
function sobre()	{
	showModalDialog('../sobre.asp',window,"font-family:Verdana; font-size:12; status:No; scrollbars:No; dialogWidth:28; dialogHeight:18")
}
function ajuda()	{
	showModalDialog('../ajuda_form_interativo.asp',window,"font-family:Verdana; font-size:12; status:No; scrollbars:No; dialogWidth:60; dialogHeight:30")
}
function cmdExec(cmd,opt) {

  	frames.message.document.execCommand(cmd,"",opt);

	frames.message.focus();

}
function tableDialog()

{
   //----- Creates A Table Dialog And Passes Values To createTable() -----
   var rtNumRows = null;
   var rtNumCols = null;
   var rtTblAlign = null;
   var rtTblWidth = null;
   var rtspacing = null;
   var rtpadding = null;
   var rtbrdsize = null;
   showModalDialog("table_news.htm",window,"status:false;dialogWidth:320px;dialogHeight:310px");
}

function createTable()
{

   //----- Criando uma tabela definida pelo usuário -----
 
   var cursor = frames.message.document.selection.createRange();
   if (rtNumRows == "" || rtNumRows == "0")
   {
      rtNumRows = "1";
   }
   if (rtNumCols == "" || rtNumCols == "0")
   {
      rtNumCols = "1";
   }
   var rttrnum=1
   var rttdnum=1
   var rtNewTable = "<table border='" + rtbrdsize + "' align='" + rtTblAlign + "' cellpadding='" + rtpadding + "' cellspacing='" + rtspacing + "' width='" + rtTblWidth + "'>"
   while (rttrnum <= rtNumRows)
   {
      rttrnum=rttrnum+1
      rtNewTable = rtNewTable + "<tr>"
      while (rttdnum <= rtNumCols)
      {
         rtNewTable = rtNewTable + "<td>&nbsp;</td>"
         rttdnum=rttdnum+1
      }
      rttdnum=1
      rtNewTable = rtNewTable + "</tr>"
   }
   rtNewTable = rtNewTable + "</table>"
   frames.message.focus();
   cursor.pasteHTML(rtNewTable);
}
function AddImage(){	
	imagePath = prompt('Digite o Endereço da Imagem', 'http://');				
	
	if ((imagePath != null) && (imagePath != "")){					
		frames.message.document.execCommand('InsertImage', false, imagePath);
  		frames.message.focus();
	}
	frames.message.focus();			
}
function Addmail(option){

	
	mail = prompt('Digite o endereço de e-mail', 'nome@provedora.com.br');
			frames.message.document.body.innerHTML =  frames.message.document.body.innerHTML + mail + '<br>'; 
	
  frames.message.focus();
	}

function doPreview(){

     temp = frames.message.document.body.innerHTML;

     preWindow= open('', 'previewWindow', 'width=500,height=440,status=yes,scrollbars=yes,resizable=yes,toolbar=no,menubar=yes');

     preWindow.document.open();

     preWindow.document.write(temp);

     preWindow.document.close();

}
function setMode(bMode) {

	var sTmp;

  	isHTMLMode = bMode;

  	if (isHTMLMode) {

		sTmp=frames.message.document.body.innerHTML;

		frames.message.document.body.innerText=sTmp;

		

	} else {

		sTmp=frames.message.document.body.innerText;

		frames.message.document.body.innerHTML=sTmp;

		

	}

  frames.message.focus();

}
function CheckForm() {
		
	  if (document.frmSendmail.titulo.value==""){
	  	msg = "           ESCOHA UM TÍTULO       \n";
		msg += "_____________________________________\n\n";
		msg += "Você deve escolher um Título para o seu classificado.\n";
		msg += "Você está sendo direcionado para o campo em branco a ser preenchido \n";
		alert(msg + "\n\n");
            frmSendmail.titulo.focus();
            return false;
      }
	  if (document.frmSendmail.cabeca.value==""){
	  	msg = "           DIGITE A CABEÇA DA MATÉRIA       \n";
		msg += "_____________________________________\n\n";
		msg += "Você deve digitar uma cabeça para a sua matéria.\n";
		msg += "A cabeça é um breve resumo ou uma chamada para a sua matéria.\n";
		msg += "Você está sendo direcionado para o campo em branco a ser preenchido \n";
		alert(msg + "\n\n");
            frmSendmail.cabeca.focus();
            return false;
      }
	  if (document.frmSendmail.autor.value==""){
	  	msg = "           DIGITE O AUTOR DA MATÉRIA       \n";
		msg += "_____________________________________\n\n";
		msg += "Você deve digitar o Autor da sua matéria.\n";
		msg += "Você está sendo direcionado para o campo em branco a ser preenchido \n";
		alert(msg + "\n\n");
            frmSendmail.autor.focus();
            return false;
      }
	  if (document.frmSendmail.mailautor.value==""){
	  	msg = "           DIGITE O E-MAIL DO AUTOR DA MATÉRIA       \n";
		msg += "_____________________________________\n\n";
		msg += "Você deve digitar o E-Mail do Autor da sua matéria.\n";
		msg += "Você está sendo direcionado para o campo em branco a ser preenchido \n";
		alert(msg + "\n\n");
            frmSendmail.mailautor.focus();
            return false;
      }
	
	if (document.frmSendmail.body.value==""){
	  	msg = "           DIGITE O TEXTO/DESCRIÇÃO DO SEU ANÚNCIO        \n";
		msg += "_____________________________________\n\n";
		msg += "Você deve digitar algum texto/descrição para cadastrar seu classificado\n";
		msg += "Você está sendo direcionado para o campo em branco a ser preenchido \n";
		alert(msg + "\n\n");
            frmSendmail.body.focus();
            return false;
      }
     
}