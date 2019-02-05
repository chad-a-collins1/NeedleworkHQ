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

<title></title>  	


</head>
<Body>
<%
'strProvider = "driver={SQL Server};server=jsc-srq-irm;database=nsas;uid=nsas;pwd=jaugustyn"	
Dim user, pass
Dim session, remove
Dim txtLoginStatus
Dim rsValidate
Dim strQuery1
Dim conn

strDBpath = Server.MapPath("\db\NWHQ.mdb")


Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

strQuery1 = "SELECT * FROM tblCategories ORDER BY cName"

'Create the recordset Object rsValidate
Set rs = Server.CreateObject("adodb.recordset")
rs.Open strQuery1, conn, 3, 3

Do While Not rs.EOF
	Response.Write rs.Fields("CategoryID") & Space(14) & rs.Fields("cName") & "<BR>"
	rs.MoveNext
Loop

rs.Close
Set rs = Nothing

conn.Close
Set conn = Nothing
%>
</Body>
</HTML>
