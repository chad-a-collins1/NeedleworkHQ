<% @LANGUAGE = "VBScript"%>
<% Response.Buffer = True %>
<html>
<head>
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<META name=VI60_DTCScriptingPlatform content="Server (ASP)">
<META name=VI60_defaultClientScript content=JavaScript>
<title></title>
</head>

<body background="paper.gif" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">

<script "language=javascript1.2">

	function fctNO() {
	window.navigate("ManageAccount?g=" + document.theForm.txtAlias.value)
	}
	
</script>	

<%'

alias = Request.QueryString("u").Item(1)
pID = Request.QueryString("u").Item(2)

strDBpath = Server.MapPath("\db\NWHQ.mdb")

Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

'This query string is used with the recordset 'rsValidate,' it selects a record from tblUserAccounts based on the UserID and Password that the user entered
strQuery = "SELECT * FROM tblConsignPatterns WHERE pID = " & pID 

Set rs = Server.CreateObject("adodb.recordset")
rs.Open strQuery, conn, 3, 3

%>

<FONT FACE="VERDANA" SIZE=2 COLOR="white">

<%

	
%>
<FORM ACTION="DiscontinueCSitem2.asp" METHOD="post" NAME="theform">
<INPUT TYPE="hidden" NAME="txtAlias" value="<%= alias %>">
<INPUT TYPE="hidden" NAME="txtPID" value="<%= pID %>">
<table border=0 bordercolor="#FFFFCC" cellpadding=0 cellspacing=0 bgcolor="lightyellow" width="100%" height="90%">
<tr>
<td>
<Center>
<%
If rs.Fields("userAlias") = alias then
	Response.Write "<FONT SIZE=5 color=" & """darkblue""" & ">Are you sure you want to delete " & rs.Fields("pName") & " ?"
	Response.Write "</center></td></TR><TR><TD><CENTER>"
	Response.Write "<INPUT TYPE=" & "'" & "SUBMIT" & "'" & "VALUE=" & "'" & "         YES         " & "'" & " >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=" & "'" & "SUBMIT" & "'" & "VALUE=" & "'" & "         NO         " & "'" & "onClick=" & "'" & "fctNO()" & "'" & ">"
Else
	Response.Write "<FONT SIZE=5 color=" & """darkblue""" & ">You are not authorized to perform this action.</FONT>"
	Response.Write "</center></td></TR><TR><TD><CENTER>"
	Response.Write "<INPUT TYPE=" & """Button""" & " VALUE=" & """         EXIT         """ & "onClick=" & """fctNO()""" & ">" 
End If
%>
</CENTER>
</TD>
</TR>
</TABLE>
</FORM>

<%

	rs.Close
	Set rs = Nothing
	
	conn.Close
	Set conn = Nothing
	
%>
<br>
<br>
<br>
</FONT>
</body>
</html>

