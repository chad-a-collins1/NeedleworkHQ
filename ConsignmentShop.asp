<html>
<head>
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<META name=VI60_DTCScriptingPlatform content="Server (ASP)">
<META name=VI60_defaultClientScript content=JavaScript>

<title>Member Services</title>  	


</head>

<body text="darkblue" style="FONT-FAMILY: Arial; FONT-SIZE: 8pt"  topMargin=0 background="yellow.jpg"> 
<STYLE TYPE="text/css">
	p {text-align:justify;font-size: 8pt;font-family: "Verdana"; }
	td {font-size: 8pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 8pt;font-family: "Arial"; color: "darkblue";}
	a {font-size: 8pt;font-family: "Baskerville Old Face"; color: "darkblue";}
</STYLE>
<%
strDBpath = Server.MapPath("/db/NWHQ.mdb")

Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

		
	strSql = "SELECT * FROM tblShops"
		
	Set rst1 = CreateObject("adodb.recordset")
	rst1.Open strSql, conn, 3, 3
	

%>

<Table width="65%" cellpadding=4 cellspacing=4>
<%
Do While Not rst1.EOF
	Response.Write "<tr><td bgcolor=A0FFFF><center>" & "<IMG SRC=" & "'" & rst1.Fields("sIcon") & "'" & " width=80 height=80>"  & "</center></td><td><a href=ConsignmentShop_2.asp?Theme=" & rst1.Fields("Theme") & "><h3>" & rst1.Fields("sName") & "</h3></a></td></tr>"


rst1.MoveNext 
Loop
%>
</Table>
</body>
</html>






















