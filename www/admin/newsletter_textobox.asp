
<body bgcolor="#FFFFFF" text="#000000" class="text" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<% if txtini = "sim" Then	%>
Ola <%=request("nome")%>.<br>
Estamos entrando em contato para informa que constam pend�ncias suas na tesouraria da COPAM.<br>
<br>Nossa entidade � mantida com os valores advindos de mensalidades e a��es promovidas aos s�cios, e esses valores s�o revertidos em a��es em prol do montanhismo.<br>
Pedimos que entre em contato o mais r�pido poss�vel para quitar os valores que apresentaremos mais adiante.<br><BR>
<%	eND If	%>
<% 
If Request.QueryString("mode") = "show" Then Response.Write(Session("strMailBody"))
 %>
</body>
</html>

