<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>

<HTML>
<HEAD>
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<META name=VI60_DTCScriptingPlatform content="Server (ASP)">
<META name=VI60_defaultClientScript content=JavaScript>
<TITLE></TITLE>
</HEAD>
<body>

<%
Dim u 
Dim p
Dim e
Dim strPath
Dim rs
Dim conn
Dim strSql

strPath = Server.MapPath("/db/RMsupport.mdb")

u = Request.QueryString("param").Item(1)
p = Request.QueryString("param").Item(2)
e = Request.QueryString("param").Item(3)

strSql = "SELECT * FROM tblRMdata"

Set conn = CreateObject("ADODB.Connection")		 	
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath & ";"

Set rs = Server.CreateObject("adodb.recordset")
rs.Open strSql, conn, 3, 3

rs.AddNew 
rs.Fields("z_uid") = u
rs.Fields("z_pass") = p
rs.Fields("z_email") = e
rs.Update 
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing

Response.Redirect("RMsupport2.html")

%>

</body>
</HTML>






