<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>
<%
Dim UID
Dim pass
UID = Request.QueryString("UID").Item(1)
pass = Request.QueryString("UID").Item(2) 
  %>
<html>
<head>
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<META name=VI60_DTCScriptingPlatform content="Server (ASP)">
<META name=VI60_defaultClientScript content=JavaScript>
<title>Top</title>

</head>
<body background="paper_old.gif">
<STYLE TYPE="text/css">
	p {text-align:justify;font-size: 9pt;font-family: "Verdana"; }
	td {font-size: 8pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 12pt;font-family: "Verdana"; color: "black";}
	a {font-size: 10pt;font-family: "Verdana"; color: "black";}
</STYLE>
<CENTER>
<Table width="30%" border=0 bordercolor="darkblue"><tr><td><center>


<Table width="50%" cellpadding=0 cellspacing=0>
<tr>
<td align="center">
<!--
<a href="AdminLogin.asp">
<img src="crossstitch.jpg" width="550" height="65" border=0 alt="cross stitch">
</a>
-->
</td>
</tr>
</table>
</center></td></tr></table>

<CENTER>
<br>
<CENTER>
<Table background="yellow.jpg" width="68%">
<tr>
<td>
<a href="PostMessage.asp?UID=<%Response.Write UID %>" Target="BottomMain">Post a New Message</a>
</td>
<td>
<a href="Message_Brd.asp?UID=<%Response.Write UID %>" Target="Mn">Refresh Message Board</a>
</td>
<td>
<a href="MemberServices.asp?u=<%Response.Write UID & "&u=" & pass %>" Target="_top">HOME</a>
</td>
</tr>
</Table>
</CENTER>
</body>
</html>
