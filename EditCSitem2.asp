<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>

<META content="text/html; charset=unicode" http-equiv=Content-Type>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<META name=VI60_DTCScriptingPlatform content="Server (ASP)">
<META name=VI60_defaultClientScript content=JavaScript>

<%

	Dim conn
	Dim rst1
	Dim strSql
	Dim strDBpath
	Dim i

pID = CInt(Request.Form("txtPID"))
alias = Request.Form("txtAlias")

strDBpath = Server.MapPath("/db/NWHQ.mdb")

Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

		
	
	strSql2 = "SELECT * FROM tblConsignPatterns WHERE pID = " & pID	
	strSql = "SELECT * FROM tblShops Where sID = " & CInt(Request.Form("txtLocation"))	
	
	Set rst2 = CreateObject("adodb.recordset")
	rst2.Open strSql2, conn, 3, 3

	Set rst = CreateObject("adodb.recordset")
	rst.Open strSql, conn, 3, 3

	


	rst2.Fields("pOwnerFName") = Request.Form("txtOwner")
	rst2.Fields("pPrice") = CCur(Request.Form("txtPrice"))
	rst2.Fields("Location") = rst.Fields("Theme")	
	rst2.Fields("pName") = Request.Form("txtTitle")
	rst2.Fields("pageCount") = Request.Form("txtCount")
	rst2.Update

	rst2.Close
	Set rst2 = Nothing

	conn.Close
	Set conn = Nothing

	Response.Redirect("ManageAccount.asp?g=" & alias)
%>




















