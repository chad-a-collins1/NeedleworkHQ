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

<title>Account Details</title>  	

</head>
<%


Dim conn
Dim userAlias
Dim strQuery
Dim strQuery2
Dim Total

userAlias = Request.QueryString("g")




%>
<body text="darkblue" link="#00ff00" vlink="#00ff00" alink="#00ff00" style="FONT-FAMILY: Arial; FONT-SIZE: 8pt"  topMargin=0 background="paper.gif"> 
<STYLE type=text/css>
	p {text-align:justify;font-size: 10pt;font-family: "Verdana"; }
	td {font-size: 10pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 10pt;font-family: "Verdana"; color: "black";}
	a {font-size: 10pt;font-family: "Verdana"; color: "darkblue";}
</STYLE>
<CENTER>
<img src="needlework.gif" width="580" height="90">
<BR><BR>

<TABLE bgcolor="lightyellow" border=0 bordercolor="silver" width="85%">
<TR><TD><CENTER><H3>Management Console, Account:&nbsp;<%= userAlias %></H3></CENTER></TD></TR>
<TR valign=top>
<TD>
</CENTER><BR><BR><CENTER>
<a href="AddNewConsignPat.asp#"><B><font color="blue">[Add a New Pattern to Sell]</font></B></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="MemberServices_tour.asp?u=guest&u=guest"><B><font color="blue">[Back to Member Services]</font></B></a><BR><BR>
<Table bgcolor="lightyellow" border=1 bordercolor="darkblue" width="80%" cellspacing=0 cellpadding=0 STYLE="table-layout:auto;border-collapse:collapse" bgcolor="WHITE" wrap=false>
<TR bgcolor="lightpink">

<TH>Edit</TH>
<TH>Thumbnail</TH>
<TH>Title</TH>
<TH>Location</TH>
<TH>Price (USD)</TH>
<TH>Discontinue</TH>
</TR>
<TR colspan="6"><TD bgcolor="lightgrey" colspan="6"><% Response.Write "Not patterns are listed for your account" %></TD></TR>

</Table>
</TD>
</TR>
<TR><TD><P>&nbsp;</P></TD></TR>
</TABLE>
</CENTER>
</body>
</html>

