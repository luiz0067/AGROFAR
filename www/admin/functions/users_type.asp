<% 
Function alertMsg

Set Rs1 = Server.CreateObject("ADODB.Recordset")
mail = "SELECT * FROM mail_interno WHERE id_para = '" & session("codigo") & "' AND lida = False ORDER BY data DESC"

Rs1.Open mail, Conn

 If Rs1.EOF AND Rs1.BOF = TRUE Then
' nada escreve
%>
  
  <%	Else
conta_msg = 0
Do Until  Rs1.EOF
conta_msg = conta_msg + 1

Rs1.MoveNext
Loop
if not Java Then
%>
  <script>
//var arr = showModalDialog("avisamail.asp",window,"font-family:Verdana; font-size:12; status:false; dialogWidth:14; dialogHeight:14" );
	
function avisa() {
		msg = "       Você tem [ <%=conta_msg%> ] Mensagen(s)  não Lida(as)!    \n";
		msg += "___________________________________________\n\n";
		//msg += "[ <%=conta_msg%> ] Mensagen(s)  não Lida(as)\n\n";
		msg += "<%=session("nome")%>, escolha uma opção:\n\n";
		msg += "Ø [OK] Para ler esta(s) mensagen(s) Agora.\n\n";
		msg += "Ø [CANCELAR] Para ler esta(s) mensagen(s) depois.\n";

if (confirm(msg + "\n")) {
location.href('mail_int_default_admin.asp');
}
else {
//volta
   }   
}
setTimeout('avisa()',1000)
</script>
<%
End If
%>
&nbsp;&nbsp;<font size="2" face="Verdana, Arial, Helvetica, sans-serif" ><img src="imagens/outros_icos/gallery/sm-alert.gif" width="14" height="14">Você 
tem <strong><font color="#FF0000"> <%=conta_msg%> </font></strong> mensagem(s) 
não lidas na sua caixa de entrada. <a href="mail_int_default_admin.asp" >Clique 
aqui para ler.</a></font> 
<%
End If
Rs1.Close
Set Rs1 = Nothing

'fim do verifica mail
End Function

Function descri_acesso
%>
Cada usu&aacute;rio possui um n&iacute;vel de acesso determinando o que 
      pode e o que n&atilde;o pode fazer.<br>
      A maioria das sess&otilde;es do sistema &eacute; de N&iacute;vel 1 (B&aacute;sico) 
      ou seja, pode navegar, ver quase tudo, mas n&atilde;o pode criar novos elementos, 
      n&atilde;o pode deletar etc.
<%
End Function

'Escreve o nome do Usuário pelo seu ID
Function PesqUser(id)

	   Set C_Reg2 = Server.CreateObject("ADODB.Recordset")
	   str_conta2 = "SELECT * FROM usuarios WHERE ID= " & id & ""
	   Set C_Reg2 = Conn.Execute(str_conta2)
  			Response.Write(C_Reg2("nome"))
		C_Reg2.Close
		Set C_Reg2=nothing
	'-------------------------------------------
End Function

Function FechaRs
	Rs.Close
	Set Rs = Nothing
End Function


'Escreve o nome do Usuário pelo seu ID
Function cidades
								Set Rs2 = Server.CreateObject("ADODB.Recordset")
			  					Sqltext = "SELECT * FROM cidades ORDER by cidades ASC"
								Rs2.Open SqlText, base_dados
							
								If Rs2.EOF AND Rs2.BOF = TRUE Then
							'nada				
									Else
								Do Until Rs2.EOF
								
								%>
                <option value="<%=Rs2("id")%>" > <%=Rs2("cidades")%></option>
                <%
								
							Rs2.MoveNext
							Loop
							End If
							Rs2.Close
							set Rs2 = nothing
End Function

Function cidadeescreve(id)
	
	   Set C_Reg2 = Server.CreateObject("ADODB.Recordset")
	   str_conta2 = "SELECT * FROM cidades WHERE ID= " & id 
	   Set C_Reg2 = Conn.Execute(str_conta2)
	
  			Response.Write(C_Reg2("cidades"))
		
		C_Reg2.Close
		Set C_Reg2=nothing
	'-------------------------------------------
End Function
Function profissao(id)
	
	   Set C_Reg2 = Server.CreateObject("ADODB.Recordset")
	   str_conta2 = "SELECT * FROM profissoes WHERE ID= " & id 
	   Set C_Reg2 = Conn.Execute(str_conta2)
	
  			Response.Write(C_Reg2("profissoes"))
		
		C_Reg2.Close
		Set C_Reg2=nothing
	'-------------------------------------------
End Function

Function veiculos_gerais(id)
	
	Set Rs2 = Server.CreateObject("ADODB.Recordset")
   		Sql="SELECT * from veiculos_gerais WHERE id= " & request("id")
   	Rs2.Open Sql, base_dados
	
	If Rs2.EOF AND Rs2.BOF = TRUE Then
	'nao existe, cria um novo registro
	Response.Write("Em Branco")
	Else
	Response.Write(Rs2("veiculo") & "-" & Rs2("marca"))
	'-------------------------------------------
	End If
End Function

Function acao_evento(id)
	
	Set Rs2 = Server.CreateObject("ADODB.Recordset")
   		Sql="SELECT * from acoes_promo WHERE id= " & request("id")
   	Rs2.Open Sql, base_dados
	
	If Rs2.EOF AND Rs2.BOF = TRUE Then
	'nao existe, cria um novo registro
	Response.Write("Em Branco")
	Else
	Response.Write(Rs2("nome_promo") & "-" & Rs2("local"))
	'-------------------------------------------
	End If
End Function
%>
