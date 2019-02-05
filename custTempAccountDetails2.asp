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
<CENTER>
<%

	Dim conn
	Dim rst1
	Dim strSql
	Dim strDBpath
	Dim i

strDBpath = Server.MapPath("/db/NWHQ.mdb")

Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

	
	cpID = CInt(Request.Form("cpOnFile"))
		
	'strSql ="SELECT * FROM tblPictures, tblCategories WHERE tblCategories.CategoryID = tblPictures.CategoryID AND tblPictures.CategoryID = " & cID 
	
		strSql = "SELECT * FROM tblCPNonMemberOrders WHERE custOrderID = " & cpID 

		
	Set rst1 = CreateObject("adodb.recordset")
	rst1.Open strSql, conn, 3, 3
		
			
%>
<Table bgcolor="lightyellow" height="600" width="750">
<TR valign="top">
<TD>
<CENTER><H4><%= rst1.Fields("custTitle") %></H4></CENTER>
<Font color="darkblue"><% 'Response.Write "<BR><BR>" & rst1.Fields("custTitle") '& " ordered on " & Left(cStr(rst1.Fields("InitDate")),10) & ". " & rst1.Fields("patWidth") & Space(1) & rst1.Fields("patUnits") & " (Width) X " & rst1.Fields("patHeight") & Space(1) & rst1.Fields("patUnits") & " (Height). " & rst1.Fields("FlossType") & " Floss Palette."  & "<BR><BR>" %></font>
<CENTER>
<a href="#"><IMG SRC="/customPatternOrders/<%= rst1.Fields("custOrderID") & ".jpg" %>" height="400" width="400" border=0></a>
<br>
<a href="#">[VIEW PATTERN]</a>
</CENTER>
</TD>
</TR>
</Table>
</body>

</html>
