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

<title>Users</title>  	


</head>
<%

Dim user, pass
Dim session, remove
Dim txtLoginStatus
Dim rsValidate
Dim strQuery1
Dim conn

strDBpath = Server.MapPath("\db\NWHQ.mdb")


Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

strQuery3 = "SELECT * FROM tblUserAccounts Order By StartDate ASC"

Set rsCustom = Server.CreateObject("adodb.recordset")
rsCustom.Open strQuery3, conn, 3, 3

Do WHile Not rsCustom.EOF
	Response.Write rsCustom.Fields("uid") & space(5) & rsCustom.Fields("pswd") & space(5) & rsCustom.Fields("Email") & "<br><br>"
	rsCustom.MoveNext
Loop
	rsCustom.Close
	Set rsCustom = Nothing	

	conn.Close
	Set conn = Nothing
%>
</body>
</html>
































