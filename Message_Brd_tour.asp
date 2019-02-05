<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>
<%
	Dim UID				'<--------------------------- THIS VALUE IS NOT GETTING PASSED IN FROM THE FRAMES PAGE
	UID = "guest"
%>
<html>
<head>
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<META name=VI60_DTCScriptingPlatform content="Server (ASP)">
<META name=VI60_defaultClientScript content=JavaScript>
<html>
<head>
<title>NWHQ Msg</title>
</head>
<body text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" bgcolor="lightyellow">
<STYLE TYPE="text/css">
	p {text-align:justify;font-size: 9pt;font-family: "Verdana"; }
	td {font-size: 8pt;font-family: "Verdana";color: "black"; }
	th {font-size: 12pt;font-family: "Verdana"; color: "black";}
	a {font-size: 10pt;font-family: "Verdana"; color: "black";}
</STYLE>

<%

	Dim strProvider
	Dim DB
	Dim strSQL
	Dim conn
	Dim rsMessage
	Dim i, row

	strDBpath = Server.MapPath("/db/NWHQ.mdb")

	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

	Set rsMessage = Server.CreateObject("adodb.recordset")

	strSQL = "SELECT * FROM tblMessageBoard ORDER BY PostDate DESC"
		
%>
<CENTER>
<FONT FACE="VERDANA" SIZE=2 COLOR="white">
<%

		rsMessage.Open strSQL, conn, 3, 3
		
		%>
		<CENTER>
		<table width="99%" border=1 bordercolor="GRAY" cellspacing=0 cellpadding=0 STYLE="table-layout:auto;border-collapse:collapse" bgcolor="WHITE" wrap=false>
		<tr>
			<th bgcolor="lightblue" size=35><center>Author</center></th>
			<th bgcolor="lightblue" size=100><center>Subject</center></th>
			<th bgcolor="lightblue" size=35><center>Date</center></th>
		</tr>
		<%
		row = 1
		Do While Not rsMessage.EOF 
			If row mod 2 = 0 Then
				%>
				<tr bgcolor="lightblue">
				<%
			Else
				%>
				<tr bgcolor="lightgrey">
				<%
			End If
			%>
				<td size=35><center><B><%= rsMessage.Fields("Author") %></B></center><input type="hidden" name="txtID" value="<%= rsMessage.Fields("ID") %>"></td>
				<td size=100><center><a href="Bottom2_tour.asp?ID=<%= rsMessage.Fields("ID") & "&ID=" & UID %>"  Target="BottomMain" ><B><FONT COLOR="blue"><%= rsMessage.Fields("Subject") %></FONT></B></a></center></td>
  				<td size=35><B><%= rsMessage.Fields("PostDate") %></B></td>			
	 		</tr>			
			<%
			rsMessage.MoveNext 
			row = row + 1
		Loop

	rsMessage.Close 
	Set rsMessage = Nothing
	conn.Close
	Set conn = Nothing
		
%>
</table>
</CENTER>
</FONT>
</body>
</html>
