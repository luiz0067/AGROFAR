<!-- #include file="../login_nivel1_limpo.asp" --><%

'seleciona o usuário para ir direto para página de edição
   	Set rs = Server.CreateObject("ADODB.Recordset")
	'Sqltext = "SELECT * FROM prospectes WHERE nome LIKE '%%" & Request("nome") & "%%' ORDER BY nome ASC"
	Sqltext = "SELECT * FROM prospectes WHERE nome = '" & Request("nome") & "' AND padrinho = '"&request("padrinho")&"' AND cidade='"&request("cidade")&"'"
	Rs.CursorLocation = 3 ' abre o bd em nível 3
	Rs.Open SqlText, base_dados ' executa a aão de busca
	id = rs("id")
   	FechaRs
	' ----------------------------------------------------
   	' Cadastra historicos e adiciona na agenda do usuários
	'-----------------------------------------------------
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open base_dados 
	hist = "ESTE PROSCPECT FOI CADASTRADO POR ESTE USUÁRIO."
		   
		Sql2="INSERT INTO prospectes_historico(usuario,id,historico,data)"
   		Sql2 = Sql2 + " VALUES('" & session("codigo") & "','" & id & "','" &hist&   "','" & now()&"')"		
	       Conn.Execute(Sql2)
'-----------------------------------
'Lógica para construir o calendário
Session.LCID = 1046 'BRASIL
if (Request.QueryString("Mes") <> Month(now))AND(Request.QueryString("Mes") <> "") Then
Mes=CInt(Request.QueryString("Mes"))'força que o resultado seja um inteiro
else
Mes = Month(now)
end if
Dia=Day(now)
Ano = Year(now)
Agora = DateSerial(Ano, mes, dia)
PrimeiroDiaMes=DateSerial(Year(Now),Mes,1)
UltimoDiaMes=DateSerial(Year(Now),Mes+1,1-1)
Inicio = ABS(1 - WeekDay(PrimeiroDiaMes))
Fim = 7 - WeekDay(UltimoDiaMes)
Start=1-Inicio
TheEnd = Day(UltimoDiaMes) + Fim 
J=1 
Agora = Agora + 1
OBS = "O Prospect <a href=""prospects_editar.asp?id="&id&""" title=""Clique aqui para editar o prospect"">" &Request("nome") & "</a> foi cadastrado no sistema, entrar em contato para agradecer a visita e atualizar dados"
TIT = "Ligar para o Prospect " & Request("nome")
'-------------------------------------
'Cadastrando no Calendário
'*************************************
	Sql3="INSERT INTO agenda (id_user,data,dia,mes,ano,obs,titulo,status,id_pro)"
   	Sql3 = Sql3 + " VALUES('" & session("codigo") & "','" & Agora & "','" & Day(Agora) & "','" & Month(Agora) & "',"
   	Sql3 = Sql3 + "'" & Year(Agora) & "','" & OBS & "','" & TIT & "','Em Andamento','"&id&"')"
   Conn.Execute(Sql3)  
%>
<%=Day(Agora)%>/<%=Month(Agora)%>/<%=Year(Agora)%>
   
   <%
   nome = Request("nome") 
   Conn.Close
Set Conn = Nothing
   %>
   <meta http-equiv="refresh" content="1;URL=../prospects_editar.asp?id=<%=id%>">
      <div align="center">
        <style>
<!--
.clsNormal    { font-family: Verdana, Arial, Helvetica, MS Sans Serif, sans-serif; font-size:12px; padding: .2em;border-color: #000000; }
.clsDynaText    { font-family: Verdana, Arial, Helvetica, MS Sans Serif, sans-serif; padding: .2em; font-size:12px; color:#ffffff; border-color: #000000;background-color:#6495ed; }
-->
</style>
        <SCRIPT LANGUAGE="JavaScript">
<!--
    var sPlainText 
    var iTick = 15
    var j = 0
    var sOffID
    
    window.onload = doOnload
    
    function doOnload() {
        idDynaText.innerHTML = replaceChars(idDynaText.innerText," ","&nbsp;")
        sPlainText = idDynaText.innerText
        getString()
    }
    
    function getString() {
        sDynaText = ""
        for (i=0;i<sPlainText.length;i++){
            sDynaText += "<SPAN ID=idDynaChar" + i + " CLASS=clsNormal>" + sPlainText.substring(i,i+1) + "</SPAN>"
        }
        idDynaText.innerHTML = sDynaText
        window.setTimeout("doDynaText()",iTick)
    }

    function doDynaText() {
        if (null != sOffID) {
        }
        sOnID = "idDynaChar" + (j)
        document.all(sOnID).className = "clsDynaText"
        sOffID = sOnID
        j++
        if (j < sPlainText.length    ) {
            window.setTimeout("doDynaText()",iTick)
        }
        else {
            if (null != sOffID) {
            }
        }
    }

    function replaceChars(sString,sOld,sNew) {
        for (i = 0; i < sString.length; i++) {
            if (sString.substring(i, i + sOld.length) == sOld) {
                sString = sString.substring(0, i) + sNew + sString.substring(i + sOld.length, sString.length)
            }
        }
        return sString
    }
-->
</SCRIPT>
 <center>
  <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Prospect 
  <font color="#FF0000"><%=Request("nome")%></font> Cadastrado com Sucesso ...<br>
  Atualizando dados, e sendo redirecionado para a página de outros cadastros no Novo Prospect.<br>
          Caso o sistema não redirecione automaticamente <a href="../prospects_editar.asp?id=<%=id%>">clique 
          aqui</a>. </strong></font><br>
          <br>
		  <br> <span id="idDynaText"  > 
          <table width="71%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              
      <td height="27">&nbsp;<font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>PROSPECT 
        CADASTRADO com Sucesso!!! Aguarde, carregando configurações....</strong></font></td>
            </tr>
          </table>
          </span> 
        </center>
		<br><br>
		<br>
		<br>