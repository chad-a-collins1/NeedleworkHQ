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
	td {font-size: 8pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 8pt;font-family: "Verdana"; color: "darkblue";}
	a {font-size: 12pt;font-family: "Verdana"; color: "blue";}
</STYLE>

<body background="paper_old.gif">

<%

	Dim conn
	Dim rst1
	Dim strSql
	Dim strDBpath
	Dim i
	Dim Count

strDBpath = Server.MapPath("/db/NWHQ.mdb")

Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

	
	pID = CInt(Request.QueryString("p"))
		
	strSql = "SELECT * FROM tblConsignPatterns WHERE pID = " & pID 
		
	Set rst1 = CreateObject("adodb.recordset")
	rst1.Open strSql, conn, 3, 3

	
	Response.Write 	"<Center><Table cellpadding=5 background='yellow.jpg' width='500'>"
	
Response.Write "<tr><th><font color='red'><h4><B>Open with PCStich Viewer</b></h4></font></th></tr>"
Response.Write "<TR><TD bgcolor=lightyellow><center><a href=dev/ConsignmentShop/" & rst1.Fields("pID") & ".pat" & "><b>" & rst1.Fields("pName") & ".PAT" & "</b></a></center><br><br></td></tr>"
Response.Write "<tr><td><center><a href='http://www.pcstitch.com/PatView/Download.ASP'><Font color='green'><B><I>Download PCStitch Viewer!</I></B></Font></a></CENTER></td></tr>"
Response.Write "</table>"
%>
</body>

</html>










