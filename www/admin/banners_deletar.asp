<!-- #include file="login_nivel3.asp" -->
<!--#include file=".\config_site.inc"-->
<%

  Sql="DELETE FROM banners WHERE B_ID=" & Request.QueryString("B_ID")
  Conn.Execute(Sql)
  Conn.Close
  Set Conn = Nothing
 
 'Response.redirect "banner_admin.asp"
 %>
 <script>alert('..::   »» BANNER DELETADO COM SUCESSO!!! ««    ::..  ');
    	location.href='banner_admin.asp';</script>
