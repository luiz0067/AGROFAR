
<%
'Administracao dos banners
Response.Buffer=True
%>

<!-- #include file="login_nivel3.asp" -->
<!--#include file=".\config_site.inc" -->
<%
dominiundados = base_dados
' *** Edit Operations: declare variables

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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then



  MM_editConnection = dominiundados
  MM_editTable = "BANNERS"
  MM_editRedirectUrl = "banners_upload.asp?B_ID=B_ID&acao=form"
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
' *** Insert Record: construct a sql insert statement and execute it

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
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
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End if
    MM_tableValues = MM_tableValues & MM_columns(i)
    MM_dbValues = MM_dbValues & FormVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close


'pega o último registro
								Sql = "SELECT Max(B_ID) AS opa FROM banners"
  								Set RS = Conn.Execute(Sql)
  								B_ID_temp =  RS("opa")
  								
  


    If (MM_editRedirectUrl <> "") Then
      Response.Redirect "banners_upload.asp?B_ID=" & B_ID_temp & "&acao=form"
    End If
  RS.Close
  Set RS = Nothing
  Set RS = Conn.Execute(Sql)
  
  
  End If

End If
%>
<%
set rsBanner = Server.CreateObject("ADODB.Recordset")
rsBanner.ActiveConnection = dominiundados
rsBanner.Source = "SELECT * FROM BANNERS ORDER BY B_ID ASC"
rsBanner.CursorType = 0
rsBanner.CursorLocation = 2
rsBanner.LockType = 3
rsBanner.Open()
rsBanner_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsBanner_numRows = rsBanner_numRows + Repeat1__numRows
%>

<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<html>
<head>
<title>Página Administrador de Banners :::::::::::Midiamix nteractive ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
   <tr bgcolor="#00CC99" valign="middle"> 
    <td align="left" height="20" colspan="2" bgcolor="#FFFFFF">
	<p align="center"><b><font color="#FFFFFF" face="Arial" size="4">&nbsp;&nbsp;</font><font color="#FFCC00" face="Arial" size="4">Banner
      Adiministração</font></b></p>
	</td>
  </tr>
      <tr align="center" valign="middle"> 
    <td valign="top" colspan="2"> 
      <form method="post" action="<%=MM_editAction%>" name="form1">
        <table align="center">
          <tr valign="baseline"> 
            <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000080">NOME:</font></b></td>
            <td> 
              <input type="text" name="B_NAME" size="56" style="background-color: #FFCC00">
            </td>
          </tr>
          <tr valign="baseline"> 
            <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000080">LINK
              URL:</font></b></td>
            <td> 
              <input type="text" name="B_URL" value="" size="56" style="background-color: #FFCC00">
            </td>
          </tr>
          
          <tr valign="baseline"> 
            <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000080">ALT:</font></b></td>
            <td> 
              <input type="text" name="B_ALT" value="" size="56" style="background-color: #FFCC00">
            </td>
          </tr>
			<tr>
      <td nowrap align="right"><font color="#000080"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>TARGET:</b></font>&nbsp;</font></td>
      <td> 
		<select size="1" name="B_TARGET" style="background-color: #FFCC00">
			<option value="_blank">Abrir nova janela</option>
			<option value="_parent">Na mesma janela</option>
		</select>
      </td>
			</tr>
          <tr valign="baseline"> 
            <td nowrap align="right"><font color="#000080">&nbsp;</font></td>
            <td align="right" valign="top"> 
			<p align="center"> <font face="Verdana, Arial, Helvetica, sans-serif"> 
              <font size="1"> <i> 
              <input type="submit" value="..::Inserir Banner::.." style="background-color: #FFCC00">
</i></font></font></p>
			</td>
          </tr>
        </table>
        <input type="hidden" name="MM_insert" value="true">
      </form>
      
    </td>
  </TR>
  
  <tr> 
    <td align="left" valign="top" colspan="2">&nbsp;</td>
  </tr>
    <tr> 
    <td align="left" valign="top" colspan="2"> 
      <table width="100%" border="0" cellspacing="5" cellpadding="2">
        <% 
While ((Repeat1__numRows <> 0) AND (NOT rsBanner.EOF)) 
%>
        <tr> 
          <td align="left" valign="top"> 
			<div align="center">
				<center> 
            <table  border="0" cellspacing="1" cellpadding="0" bgcolor="#FFCC00" width="703">
              <tr align="left" valign="top" bgcolor="#FFFFFF"> 
                <td valign="middle" align="center" width="166" ><b>
                <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#0000FF">
                <%=(rsBanner.Fields.Item("B_NAME").Value)%></font></b> <br>
                [ <a href="banner_update.asp?id=<%=(rsBanner.Fields.Item("B_ID").Value)%>"><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Editar</font></a>
				]<font
				face="Verdana, Arial, Helvetica, sans-serif" size="1"> - [ <a href="banners_deletar.asp?B_ID=<%=(rsBanner.Fields.Item("B_ID").Value)%>">Deletar</a>] </font></td>
                <td  valign="middle" align="center" bgcolor="#FFCC00" width="53"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#0000FF"><%=(rsBanner.Fields.Item("B_CLICKED_TOTAL").Value)%> </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000"> cliques
                  </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#0000FF"><br>
                      </font></b><font color="#666666" size="1"><i><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000">(<%=(rsBanner.Fields.Item("B_ADDED_DATE").Value)%> 
                      - <%=(rsBanner.Fields.Item("B_CLICKED_DATE").Value)%>)</font></i></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000"><br>
                      ID:<%=(rsBanner.Fields.Item("B_ID").Value)%> &nbsp; </font></td>
                    <td align="right" width="470" bgcolor="#FFCC00"> <a href="../banner_redirect.asp?id=<%=(rsBanner.Fields.Item("B_ID").Value)%>&url=<%=(rsBanner.Fields.Item("B_URL").Value)%>" target="<%=(rsBanner.Fields.Item("B_TARGET").Value)%>"> 
                      <font size="1" face="Verdana, Arial, Helvetica, sans-serif">[Clique 
                      aqui para testar o Banner] </font>
                      <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="480" height="70">
                        <param name="movie" value="../<%=(rsBanner.Fields.Item("B_IMAGE").Value)%>">
                        <param name="quality" value="high">
                        <embed src="../<%=(rsBanner.Fields.Item("B_IMAGE").Value)%>" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="480" height="70"></embed> 
                      </object>
				</a></td>
              </tr>
            </table>
				</center>
			</div>
          </td>
        </tr>
        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsBanner.MoveNext()
Wend
%>
      </table>
    </td>
  </tr>
  
</table>
</body>
</html>
<%
rsBanner.Close()




%>


