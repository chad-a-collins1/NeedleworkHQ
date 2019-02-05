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

<title>Add Consignment Pattern</title>  	
</head>
<%
Dim pID
Dim strQuery
Dim conn

pID = Request.QueryString("p")

strDBpath = Server.MapPath("\db\NWHQ.mdb")

Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

strQuery = "SELECT * FROM tblConsignPatterns WHERE pID = " & CInt(pID)


Set rs = Server.CreateObject("adodb.recordset")
rs.Open strQuery, conn, 3, 3
%>

<BODY background="paper_old.gif">
<img src="dev/ConsignmentShop/<% Response.Write pID & "_pat.jpg" %>">
</BODY>
</html>