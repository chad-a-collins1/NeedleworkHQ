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

<title>verify cust</title>  	


</head>
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

'Set the UserID and Password variables equal to the user input via the txtUserName and txtPassword text fields, respectively.
user = Request.Form("txtEmail")
pass = Request.Form("txtPass")

	strQuery1 = "SELECT * FROM tblCPNonMembers WHERE Email = " & "'" & user & "'" & " AND Pass = " & "'" & pass & "'" 

'Create the recordset Object rsValidate
Set rsValidate = Server.CreateObject("adodb.recordset")
rsValidate.Open strQuery1, conn, 3, 3


If Not rsValidate.BOF And Not rsValidate.EOF Then
	Response.Redirect("custTempAccountDetails.asp?u=" & rsValidate.Fields("custCustomerID"))	

ElseIf rsValidate.EOF Then
	Response.Redirect("LoginFailed.asp?")
		
	rsValidate.Close
	Set rsValidate = Nothing
	
	conn.Close
	Set conn = Nothing
	
End If
'session.Abandon 
%>

</body>
</html>
































