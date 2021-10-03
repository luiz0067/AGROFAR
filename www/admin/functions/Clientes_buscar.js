function abre_rela_print()	{
	window.open('prospect_relatorio_print.asp?id=<%=id%>&texto=<%=request("texto")%>','WinExa','toolbar=1,location=0,status=0,menubar=0,scrollbars=1,resizable=1,width=800,height=600')
}
function abre_rela_pdf()	{
	window.open('prospect_relatorio_print_pdf.asp?texto=<%=request("texto")%>','WinExa','toolbar=1,location=1,status=0,menubar=1,scrollbars=0,resizable=0,width=800,height=600')
}
function m_apo(){

var message="Prospect Aposentado!"
var neonbasecolor="#0000ff"
var neontextcolor="red"
var flashspeed=100  //in milliseconds

///No need to edit below this line/////

var n=0
if (document.all){
document.write('<font color="'+neonbasecolor+'">')
for (m=0;m<message.length;m++)
document.write('<span id="neonlight">'+message.charAt(m)+'</span>')
document.write('</font>')

//cache reference to neonlight array
var tempref=document.all.neonlight
}
else
document.write(message)

function neon(){

//Change all letters to base color
if (n==0){
for (m=0;m<message.length;m++)
tempref[m].style.color=neonbasecolor
}

//cycle through and change individual letters to neon color
tempref[n].style.color=neontextcolor

if (n<tempref.length-1)
n++
else{
n=0
clearInterval(flashing)
setTimeout("beginneon()",1500)
return
}
}

function beginneon(){
if (document.all)
flashing=setInterval("neon()",flashspeed)
}
beginneon()

}

var version4 = (navigator.appVersion.charAt(0) == "4"); 
var popupHandle;
function closePopup() {
if(popupHandle != null && !popupHandle.closed) popupHandle.close();
}
function displayPopup(position,url,name,height,width,evnt) {
var properties = "toolbar = 0, location = 0, height = " + height;
properties = properties + ", width=" + width;
var leftprop, topprop, screenX, screenY, cursorX, cursorY, padAmt;
if(navigator.appName == "Microsoft Internet Explorer") {
screenY = document.body.offsetHeight;
screenX = window.screen.availWidth;
}
else {
screenY = window.outerHeight
screenX = window.outerWidth
}
if(position == 1)	{ 
cursorX = evnt.screenX;
cursorY = evnt.screenY;
padAmtX = 10;
padAmtY = 10;
if((cursorY + height + padAmtY) > screenY) {

padAmtY = (-30) + (height * -1);

}
if((cursorX + width + padAmtX) > screenX)	{
padAmtX = (-30) + (width * -1);	

}
if(navigator.appName == "Microsoft Internet Explorer") {
leftprop = cursorX + padAmtX;
topprop = cursorY + padAmtY;
}
else {
leftprop = (cursorX - pageXOffset + padAmtX);
topprop = (cursorY - pageYOffset + padAmtY);
   }
}
else{
leftvar = (screenX - width) / 2;
rightvar = (screenY - height) / 2;
if(navigator.appName == "Microsoft Internet Explorer") {
leftprop = leftvar;
topprop = rightvar;
}
else {
leftprop = (leftvar - pageXOffset);
topprop = (rightvar - pageYOffset);
   }
}
if(evnt != null) {
properties = properties + ", left = " + leftprop;
properties = properties + ", top = " + topprop;
}
closePopup();
popupHandle = open(url,name,properties);
}
