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
<script language="javascript1.2">

function fctCancel(u) {

	window.navigate("ManageAccount.asp?g="+u)
}

</script>

</head>

<STYLE TYPE="text/css">
	p {text-align:justify;font-size: 9pt;font-family: "Verdana"; }
	td {font-size: 10pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 12pt;font-family: "Verdana"; color: "black";}
	a {font-size: 8pt;font-family: "Verdana"; color: "blue";}
</STYLE>

<body background="paper.gif">

<%

	Dim conn
	Dim rst1
	Dim strSql
	Dim strDBpath
	Dim i

If Request.QueryString("g").COunt = 2 Then

	pID = CInt(Request.QueryString("g").Item(1))
	alias = Request.QueryString("g").Item(2)

strDBpath = Server.MapPath("/db/NWHQ.mdb")

Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

		
	
	strSql2 = "SELECT * FROM tblConsignPatterns WHERE pID = " & pID	
	strSql = "SELECT * FROM tblShops"
	
	Set rst2 = CreateObject("adodb.recordset")
	rst2.Open strSql2, conn, 3, 3

	Set rsShop = CreateObject("adodb.recordset")
	rsShop.Open strSql, conn, 3, 3

%>

<br><br><br><br>
<center>
<Form action="EditCSitem2.asp" Method="post" name="theForm">
<table bgcolor="lightyellow" width="40%" height="350">
<tr colspan=2><th colspan=2><center><B>Edit Console: Pattern #<%= pID %></B></center></th></tr> 
<tr><td><b>Displayed Owner Name:</b></td><td><input type="text" name="txtOwner" value="<%= rst2.Fields("pOwnerFName") %>"</td></tr>
<tr><td><b>Price:</b></td><td><input type="text" name="txtPrice" value="<%= rst2.Fields("pPrice") %>"</td></tr>
<tr><td><b>Page COunt:</b></td><td><input type="text" name="txtCount" value="<%= rst2.Fields("pageCount") %>"</td></tr>
<tr><td><b>Displayed Title:</b></td><td><input type="text" name="txtTitle" value="<%= rst2.Fields("pName") %>"</td></tr>
<tr><td><b>Shop Location:</b></td>
<td>
<select name="txtLocation" selected="<%= rst2.Fields("Location") %>">

<%
Do While Not rsShop.EOF

Response.Write "<option value=" & rsShop.Fields("sID") & ">" & rsShop.Fields("sName") & "</option>"

rsShop.MoveNext
Loop
%>

</select>
</td>
</tr>
<tr><td><Input Type="hidden" name="txtAlias" value="<%= alias %>"></td><td><Input Type="hidden" name="txtPID" value="<%= pID %>"></td></tr>
<tr><td><Input type="button" value=" Cancel " onClick="fctCancel(txtAlias.value)"></td><td><input type="submit" value="  Submit  "></td></tr>
</table>
</form>
</center>
<%
rst2.CLose
Set rst2 = Nothing

conn.Close
Set conn = Nothing

Else
	Response.Write "Transfer Failure"
End If

%>
</body>
</html>










