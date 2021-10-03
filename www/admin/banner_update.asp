
<!-- #include file="login_nivel3.asp" -->
<!--#include file=".\config_site.inc" -->
<%
dominiundados = base_dados
' *** Edit Operations: declare variables

ID_TEMP = Request.QueryString("ID")

MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = dominiundados
  MM_editTable = "BANNERS"
  MM_editColumn = "B_ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "banner_admin.asp?B_ID=" & ID_TEMP
  MM_fieldsStr  = "B_NAME|value|B_URL|value|B_ALT|value|B_TARGET|value"
  MM_columnsStr = "B_NAME|',none,''|B_URL|',none,''|B_ALT|',none,''|B_TARGET|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(i+1) = CStr(Request.Form(MM_fields(i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    FormVal = MM_fields(i+1)
    MM_typeArray = Split(MM_columns(i+1),",")
    Delim = MM_typeArray(0)
    If (Delim = "none") Then Delim = ""
    AltVal = MM_typeArray(1)
    If (AltVal = "none") Then AltVal = ""
    EmptyVal = MM_typeArray(2)
    If (EmptyVal = "none") Then EmptyVal = ""
    If (FormVal = "") Then
      FormVal = EmptyVal
    Else
      If (AltVal <> "") Then
        FormVal = AltVal
      ElseIf (Delim = "'") Then  ' escape quotes
        FormVal = "'" & Replace(FormVal,"'","''") & "'"
      Else
        FormVal = Delim + FormVal + Delim
      End If
    End If
    If (i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(i) & " = " & FormVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      'Response.Redirect(MM_editRedirectUrl)
     'Response.Redirect "banners_upload.asp?acao=form&B_ID=" & Request.QueryString("ID")   
    %>
    	<script>alert('..:: Banner atualizado com Exito - Chamando o Upload da Imagem do Banner ::..');
    	location.href='banners_upload.asp?acao=form&B_ID=<%=Request.QueryString("ID")%>';</script>
			<%
    
    End If
  End If

End If
%>
<%
Dim rsBanner__MMColParam
rsBanner__MMColParam = "1"
if (Request.QueryString("id") <> "") then rsBanner__MMColParam = Request.QueryString("id")
%>
<%
set rsBanner = Server.CreateObject("ADODB.Recordset")
rsBanner.ActiveConnection = dominiundados
rsBanner.Source = "SELECT * FROM BANNERS WHERE B_ID = " + Replace(rsBanner__MMColParam, "'", "''") + ""
rsBanner.CursorType = 0
rsBanner.CursorLocation = 2
rsBanner.LockType = 3
rsBanner.Open()
rsBanner_numRows = 0
%>
<html>
<head>
<title>Banner Atualizar Dados</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<form method="POST" action="<%=MM_editAction%>" name="form1">
  <table align="center" cellspacing="4" cellpadding="2">
    <tr valign="baseline"> 
      <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">NOME:</font></b></td>
      <td> 
        <input type="text" name="B_NAME" value="<%=(rsBanner.Fields.Item("B_NAME").Value)%>" size="32" style="background-color: #FFCC00">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">LINK
        URL:</font></b></td>
      <td> 
        <input type="text" name="B_URL" value="<%=(rsBanner.Fields.Item("B_URL").Value)%>" size="32" style="background-color: #FFCC00">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">ALT:</font></b></td>
      <td> 
        <input type="text" name="B_ALT" value="<%=(rsBanner.Fields.Item("B_ALT").Value)%>" size="32" style="background-color: #FFCC00">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>TARGET:</b></font>&nbsp;</td>
      <td> 
		<select size="1" name="B_TARGET" style="background-color: #FFCC00">
			<option value="_blank">Abrir nova janela</option>
			<option value="_parent">Na mesma janela</option>
		</select>
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"></td>
      <td> 
        <input type="submit" value="Atulizar Banner" style="background-color: #FFCC00">
      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_update" value="true">
  <input type="hidden" name="MM_recordId" value="<%= rsBanner.Fields.Item("B_ID").Value %>">
</form>
<p>&nbsp;</p>
</body>
</html>
<%
rsBanner.Close()
%>

