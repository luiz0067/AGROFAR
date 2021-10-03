<%
Casa = True
Net  = False
'---------------------
'Casa = False
'Net  = True
'---------------------
drive_casa = "e"


'##### Na Internet #####

if Net Then

'#### Caminho Base de Dados

	'escalada = "DBQ=ehome/agrofar/Dados/agrofar.mdb;Driver={Microsoft Access Driver (*.mdb)};"
	escalada = drive_casa&":home/agrofar/Dados/agrofar.mdb;"

'#### UPLOAD BD

	caminho_upload_clientes = drive_casa&":home/agrofar/Dados"

'#### ARQUIVO TXT ATUALIZA BD

	arquivo_txt_atualiza_bd= drive_casa&":home/agrofar/Dados/clientes.txt"

'#### Upload de Fotos Usuários

	usuario_dir = drive_casa&":/home/agrofar/web/upload/usuarios"
	logo_dir = drive_casa&":/home/agrofar/web/upload/fotos"
	usados_dir = drive_casa&":/home/agrofar/web/upload/usados"


'##### Em Casa #####
elseif Casa Then

'#### Caminho Base de Dados

	'escalada = "DBQ=" & drive_casa & ":home/agrofar/Dados/agrofar.mdb;Driver={Microsoft Access Driver (*.mdb)};"
	escalada =  drive_casa &":home/agrofar/Dados/agrofar.mdb;"

'#### UPLOAD BD

	caminho_upload_clientes = drive_casa & ":home/agrofar/Dados"

'#### ARQUIVO TXT ATUALIZA BD

	'arquivo_txt_atualiza_bd= drive_casa & ":home/agrofar/Dados/clientes.txt"

'#### Uploads

	usuario_dir = drive_casa & ":/home/agrofar/web/upload/usuarios"
	logo_dir = drive_casa&":/home/agrofar/web/upload/fotos"
	usados_dir = drive_casa & ":/home/agrofar/web/upload/usados"


End If

'Set Cnn = Server.CreateObject("ADODB.Connection") 
'Cnn.open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=c:/home/agrofar/web/mydatabase.mdb" 

 Set Conn = Server.CreateObject("ADODB.Connection")
 Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & escalada 


base_Dados = Conn

'------------------------------------------------
' Configurações do Site
set Rs = server.createobject("ADODB.recordset")
Sqltext = "SELECT * FROM configuracoes_site"
Rs.Open SqlText, Conn
	
cor_fundo_site = rs("cor_fundo_site")
mail_int_n = rs("mail_int_n")
host_mail = rs("host_mail")	
site_titulo = rs("site_titulo") 
site_end = rs("site_end")  
admin_mail = rs("admin_mail")  
desenvolvedor = rs("desenvolvedor")
desenvolvedor_site = rs("desenvolvedor_site")  
logo_empresa = "../html/images/logo_agro.jpg"'rs("logo_empresa")  
copy = rs("copy")  
Dados = rs("Dados")
cor_fundo_box  = rs("cor_fundo_box")
cor_texto_box  = rs("cor_texto_box")
tamanho_fonte_box  = rs("tamanho_fonte_box")
tamanho_box  = rs("tamanho_box")
cor_fonte_botao  = rs("cor_fonte_botao")
cor_fundo_botao  = rs("cor_fundo_botao")
tamanho_fonte_botao  = rs("tamanho_fonte_botao")
tamanho_botao  = rs("tamanho_botao")
cor_texto = rs("cor_texto")
cor_tabela = rs("cor_tabela")
	

' barra de rolamentos
' ---------------------------------
bar_face = rs("bar_face")
bar_shadow = rs("bar_shadow")
bar_seta = rs("bar_seta")
bar_luz = rs("bar_luz")
bar_linha_direita = rs("bar_linha_direita")
bar_3d = rs("bar_3d")
cor_links = rs("cor_links")
site_fonte = rs("site_fonte")
' fim barra de rolamentos
' ---------------------------------	
	%>
	<style>
     a {  text-decoration: none; color: <%=cor_links%>}
a:hover {  text-decoration: underline}
 
	body {
	scrollbar-face-color:<%=bar_face%>;
	scrollbar-shadow-color:<%=bar_shadow%>;
	scrollbar-highlight-color:<%=bar_luz%>;
	scrollbar-3dlight-color:<%=bar_3d%>;
	scrollbar-darkshadow-color:<%=bar_linha_direita%>;
	scrollbar-arrow-color:<%=bar_seta%>;
	}
  </style>
	<%
	
	rs.Close
	'Connexao.Close
	Set Rs = Nothing
	'Set Connexao = Nothing

'**************************
'caminhos do upload
	
	

'***************************

%>
<!--#include file="functions/users_type.asp"-->



<link href="functions/class.css" type="text/css" rel="stylesheet" />

