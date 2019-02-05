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

<title>custTempAccountDetails1</title>  	


</head>
<%
'strProvider = "driver={SQL Server};server=jsc-srq-irm;database=nsas;uid=nsas;pwd=jaugustyn"	
Dim user, pass
Dim rsValidate
Dim strQuery9
Dim conn

strDBpath = Server.MapPath("\db\NWHQ.mdb")

Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

'Set the UserID and Password variables equal to the user input via the txtUserName and txtPassword text fields, respectively.
user = Request.Form("txtEmail")
pass = Request.Form("txtPass")

	strQuery9 = "SELECT * FROM tblCPNonMembers WHERE Email = " & "'" & user & "'" & " AND Pass = " & "'" & pass & "'"

	Set rsValidate = Server.CreateObject("adodb.recordset")
	rsValidate.Open strQuery9, conn, 3, 3
	

If Not rsValidate.BOF And Not rsValidate.EOF Then
	
	strQuery3 = "SELECT * FROM tblCPNonMemberOrders WHERE custCustomerID = " & rsValidate.Fields("custCustomerID")
	
	Set rsCustom = Server.CreateObject("adodb.recordset")
	rsCustom.Open strQuery3, conn, 3, 3

ElseIf rsValidate.EOF Then
	Response.Redirect("LoginFailed.asp?")
		
	rsValidate.Close
	Set rsValidate = Nothing
	
End If


%>

<STYLE TYPE="text/css">
	p {text-align:justify;font-size: 9pt;font-family: "Verdana"; }
	td {font-size: 10pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 8pt;font-family: "Verdana"; color: "darkblue";}
	a {font-size: 8pt;font-family: "Verdana"; color: "blue";}
</STYLE>


</head>
<body background="paper.gif">
<CENTER>
<img src="needlework.gif" width="580" height="90">
</CENTER>
<BR>
<BR>
<BR>
<CENTER>

<Form Action="custTempAccountDetails2.asp" Method="Post" Name="theForm" ><CENTER>
<Table bgcolor="lightyellow" height="350" width="700" border=0 cellpadding=5 cellspacing=5>
<TR valign="top"><TD><B><H3>The list below contains all of the patterns you have ordered. Please select a pattern from the list and click the "VIEW" button.</H3></B>
<BR><BR><BR><CENTER>
<SELECT Size=7 Name="cpOnFile" style="WIDTH: 160px">
<%	
Dim j
Dim cost

j = 0

	Do While Not rsCustom.EOF
			Response.Write "<OPTION VALUE=' " & rsCustom.Fields("custOrderID") & "'>"
			Response.Write rsCustom.Fields("custTitle")
			Response.Write "</OPTION>"
			If rsCustom.Fields("ReadyYN") = 1 Then
				j = j + 1
			End If
			rsCustom.MoveNext 
	Loop
	
	cost = j * 5
	
	rsCustom.Close
	Set rsCustom = Nothing
	

%>	
</SELECT>
<br><br>
<input type="submit" value="View" name="view">
</CENTER>
</TD>
</TR>
</Table>
</Form>
</body>

</html>
