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


		strSql = "SELECT * FROM tblCategories"

		
	Set rsCP = CreateObject("adodb.recordset")
	rsCP.Open strSql, conn, 3, 3
		

rsCP.AddNew 
rsCP.Fields("cName") = "Music"

rsCP.Update 

rsCP.Close
Set rsCP = nothing

conn.Close
Set conn = nothing


%>

</body>
</html>











