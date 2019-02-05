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

	
	u = Request.Form("txtEmail")
	p = Request.Form("txtPass")	
	'strSql ="SELECT * FROM tblPictures, tblCategories WHERE tblCategories.CategoryID = tblPictures.CategoryID AND tblPictures.CategoryID = " & cID 
	
strSql1 = "SELECT * FROM tblCPNonMembers WHERE Email = " & "'" & u & "'"
Set rsUser = CreateObject("adodb.recordset")
rsUser.Open strSql1, conn, 3, 3
		
If Not rsUser.BOF and Not rsUser.EOF Then

		strSql2 = "SELECT * FROM tblCPNonMemberOrders WHERE custCustomerID = " & CInt(rsUser.Fields("custCustomerID"))
		Set rsCP = CreateObject("adodb.recordset")
		rsCP.Open strSql2, conn, 3, 3

		rsCP.AddNew 
		rsCP.Fields("custTitle") = Request.Form("txtCPname")
		rsCP.Fields("custCustomerID") = rsUser.Fields("custCustomerID")
		rsCP.Fields("custOrderDate") = CDate(Now())
		rsCP.Fields("custDesiredH") = 300
		rsCP.Fields("custDesirerdW") = 300
		'rsCP.Fields("FlossType") = Request.Form("txtFloss")
		'rsCP.Fields("ReadyYN") = 0		
		rsCP.Update 

Else
		rsUser.AddNew
		rsUser.Fields("Email") = Request.Form("txtEmail")
		rsUser.Fields("Pass") = Request.Form("txtPass")
		rsUser.Fields("FName") = Request.Form("txtFName")
		rsUser.Fields("LName") = Request.Form("txtLName")
		rsUser.Fields("Address1") = Request.Form("txtAddress1")
		rsUser.Fields("Address2") = Request.Form("txtAddress2")
		rsUser.Fields("City") = Request.Form("txtCity")
		rsUser.Fields("State") = Request.Form("txtState")
		rsUser.Fields("ZIP") = Request.Form("txtZIP")
		rsUser.Fields("Country") = Request.Form("txtCountry")
		rsUser.Update 
		rsUser.Close
		Set rsUser = Nothing
				
		strSql3 = "SELECT * FROM tblCPNonMembers WHERE Email = " & "'" & u & "'"
		Set rsUser2 = CreateObject("adodb.recordset")
		rsUser2.Open strSql3, conn, 3, 3				
				
		strSql2 = "SELECT * FROM tblCPNonMemberOrders WHERE custCustomerID = " & rsUser2.Fields("custCustomerID")
		Set rsCP = CreateObject("adodb.recordset")
		rsCP.Open strSql2, conn, 3, 3

		rsCP.AddNew 
		rsCP.Fields("custTitle") = Request.Form("txtCPname")
		rsCP.Fields("custCustomerID") = rsUser2.Fields("custCustomerID")
		rsCP.Fields("custOrderDate") = CDate(Now())
		rsCP.Fields("custDesiredH") = 300
		rsCP.Fields("custDesirerdW") = 300
		rsCP.Fields("Description") =Request.Form("Description")
		'rsCP.Fields("FlossType") = Request.Form("txtFloss")
		'rsCP.Fields("ReadyYN") = 0		
		rsCP.Update 
		
		rsUser2.Close 
		Set rsUser2 = Nothing

End If

rsCP.Close
Set rsCP = nothing

conn.Close
Set conn = nothing

'Response.Redirect("MemberServices.asp?u=" & "'" & u & "'&u=" & "'" & p & "'")
Response.Redirect("EmailPicture.asp?u=" & "'" & u & "'&u=" & "'" & p & "'")
%>

</body>
</html>











