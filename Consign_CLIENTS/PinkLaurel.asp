<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>


<html>
<head>
<title>NSAS Schedule</title></head>
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<META name=VI60_DTCScriptingPlatform content="Server (ASP)">
<META name=VI60_defaultClientScript content=JavaScript>

<STYLE TYPE="text/css">
	p {text-align:justify;font-size: 9pt;font-family: "Verdana"; }
	td {font-size: 10pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 8pt;font-family: "Verdana"; color: "darkblue";}
	a {font-size: 8pt;font-family: "Verdana"; color: "blue";}
</STYLE>

<body background="paper.gif">

<%

	Dim conn
	Dim rst1
	Dim strSql
	Dim strDBpath
	Dim i

strDBpath = Server.MapPath("/db/NWHQ.mdb")

Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

		
	
	strSql2 = "SELECT * FROM tblConsignPatterns" ' WHERE pID = " & CInt(32)		

	
	Set rst2 = CreateObject("adodb.recordset")
	rst2.Open strSql2, conn, 2, 3

'rst2.Fields("ActiveYN") = CBool(1)
'rst2.Update

Do While Not rst2.EOF
Response.Write rst2.Fields("pID") & Space(8) & rst2.Fields("pName") & "<BR><BR>"
rst2.MoveNext
Loop

rst2.CLose
Set rst2 = Nothing

conn.Close
Set conn = Nothing

%>
</body>
</html>










