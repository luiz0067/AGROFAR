
<body bgcolor="#FFFFFF" text="#000000"  leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<font size="1" face="Verdana, Arial, Helvetica, sans-serif">
<% 
	
If Request.QueryString("mode") = "show" Then Response.Write(Session("texto"))

 %>
</body>
</html>

