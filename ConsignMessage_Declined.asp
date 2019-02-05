
<% 
uid = Request.QueryString("u").Item(1)
pass = Request.QueryString("u").Item(2) 
%>
<html>
<body>
<STYLE type=text/css>
	p {font-size: 8pt;font-family: "Verdana";color: "darkblue"; }
	td {font-size: 10pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 8pt;font-family: "Verdana"; color: "black";}
	a {font-size: 10pt;font-family: "Verdana"; color: "blue";}
</STYLE>
<br><br><br><br><center>
<table background="yellow.jpg" width="50%">
<tr><td><center><h2><b>Your consignment pattern order has been canceled.</b></h2></center></td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td><center><h2><b><a href="MemberServices.asp?u=<%= uid & "&u=" & pass %>">[Return to Member Services]</b></h2></a></center></td></tr>
</table>
</body>
</html>






