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

<body>
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
	
		strSql = "SELECT * FROM tblCPmembers WHERE cpID = " & cpID 

		
	Set rst1 = CreateObject("adodb.recordset")
	rst1.Open strSql, conn, 3, 3
		
			
%>
<Table background="yellow.jpg" height="600" width="750">
<TR valign="top">
<TD>
<CENTER><H5><%= rst1.Fields("cpName") %></H5></CENTER>
<center><Font color="darkblue"><% Response.Write rst1.Fields("cpName") & " ordered on " & Left(cStr(rst1.Fields("InitDate")),10) & ". " & rst1.Fields("patWidth") & Space(1) & rst1.Fields("patUnits") & " (Width) X " & rst1.Fields("patHeight") & Space(1) & rst1.Fields("patUnits") & " (Height). " & rst1.Fields("FlossType") & " Floss Palette."  & "<BR><BR>" %></font>
</center>
<CENTER>
<a href="#"><IMG SRC="/dev/MembersCustomPatterns/<%= rst1.Fields("cpName") & ".jpg" %>" height="<%= rst1.Fields("imgHeight") %>" width="<%= rst1.Fields("imgWidth") %>" border=0></a>
<br>
<a href="dev/MembersCustomPatterns/<%= rst1.Fields("cpName") & ".jpg" %> ">[View Virtual Stitched Picture (enlarged)]</a><br><br>
<a href="dev/MembersCustomPatterns/<%= rst1.Fields("cpName") & ".PAT" %> ">[Download Pattern File (PC Stitch Viewer Format]</a>
</CENTER>
</TD>
</TR>
</Table>
</body>

</html>
