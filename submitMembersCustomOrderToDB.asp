<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>


<html>
<head>
<title>submit Custom Order</title></head>
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

	
	u = Request.Form("txtUID")
	p = Request.Form("txtPassword")	
	'strSql ="SELECT * FROM tblPictures, tblCategories WHERE tblCategories.CategoryID = tblPictures.CategoryID AND tblPictures.CategoryID = " & cID 
	
strSql1 = "SELECT * FROM tblUserAccounts WHERE uid = " & "'" & u & "'"
Set rsUser = CreateObject("adodb.recordset")
rsUser.Open strSql1, conn, 3, 3
		
If Not rsUser.BOF and Not rsUser.EOF Then

		strSql2 = "SELECT * FROM tblCPMembers WHERE uID = " & "'" & rsUser.Fields("uid") & "'"
		Set rsCP = CreateObject("adodb.recordset")
		rsCP.Open strSql2, conn, 3, 3

		rsCP.AddNew 
		rsCP.Fields("cpName") = Request.Form("txtCPname")
		rsCP.Fields("uID") = rsUser.Fields("uid")
		rsCP.Fields("InitDate") = CDate(Now())
		rsCP.Fields("imgHeight") = 300
		rsCP.Fields("imgWidth") = 300
		rsCP.Fields("patWidth") = Request.Form("txtPatWidth")
		rsCP.Fields("patHeight") = Request.Form("txtPatHeight")
		rsCP.Fields("FlossType") = Request.Form("txtFloss")
		rsCP.Fields("ReadyYN") = 0		
		rsCP.Fields("Description") = Request.Form("Description")
		rsCP.Update 

Else
			
		Response.Write("There is an error somewhere on this page Chad.")

End If

rsCP.Close
Set rsCP = nothing

conn.Close
Set conn = nothing

'Response.Redirect("MemberServices.asp?u=" & "'" & u & "'&u=" & "'" & p & "'")
Response.Redirect("EmailPicture_Members.asp?u=" & "'" & u & "'&u=" & "'" & p & "'")
%>

</body>
</html>











