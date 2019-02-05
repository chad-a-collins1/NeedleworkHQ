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
<% 
uid = Request.QueryString("u").Item(1)
pass = Request.QueryString("u").Item(2) 
pID = Request.QueryString("u").Item(3) 

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRs2 = Server.CreateObject("ADODB.Recordset")

oConn.Open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("./") & "\db\NWHQ.mdb"

sSQL2 = "SELECT pID, PaymentReceivedYN From tblConsignPatterns WHERE pID = " & pID
oRs2.Open sSQL2, oConn, 3, 3

oRs2.Fields("PaymentReceivedYN") = CBool(1)
oRs2.Update 
oRs2.Close 
Set oRs2 = Nothing
oConn.Close
Set oConn = Nothing
%>

<body>
<STYLE type=text/css>
	p {font-size: 8pt;font-family: "Verdana";color: "darkblue"; }
	td {font-size: 10pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 8pt;font-family: "Verdana"; color: "black";}
	a {font-size: 10pt;font-family: "Verdana"; color: "blue";}
</STYLE>
<br><br><br><br><center>
<table background="yellow.jpg" width="50%">
<tr><td><center><h2><b>Your consignment pattern has been activated.</b></h2></center></td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td><center><h2><b><a href="MemberServices.asp?u=<%= uid & "&u=" & pass %>">[Return to Member Services]</b></h2></a></center></td></tr>
</table>
</body>
</html>






