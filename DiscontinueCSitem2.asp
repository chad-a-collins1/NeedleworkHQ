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

alias = Request.Form("txtAlias")
pID = Request.Form("txtpID")

strDBpath = Server.MapPath("\db\NWHQ.mdb")

Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

'This query string is used with the recordset 'rsValidate,' it selects a record from tblUserAccounts based on the UserID and Password that the user entered
strQuery = "SELECT * FROM tblConsignPatterns WHERE pID = " & pID 

Set rs = Server.CreateObject("adodb.recordset")
rs.Open strQuery, conn, 3, 3

rs.Fields("ActiveYN") = CBool(0)
rs.Update 

rs.Close
Set rs = Nothing
	
conn.Close
Set conn = Nothing

Response.Redirect("ManageAccount.asp?g=" & alias)
%>

<FONT FACE="VERDANA" SIZE=2 COLOR="white">

<%

	
%>
<FORM ACTION="DiscontinueCSitem2.asp" METHOD="post" NAME="theform">
<INPUT TYPE="hidden" NAME="txtAlias" value="<%= alias %>">
<INPUT TYPE="hidden" NAME="txtPID" value="<%= pID %>">

</FORM>
<br>
<br>
<br>
</FONT>
</body>
</html>

