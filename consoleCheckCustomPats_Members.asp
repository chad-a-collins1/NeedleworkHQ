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

<body bgcolor="darkblue">
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

If Request.QueryString("uID").Count = 0 Then	
	u = Request.Form("txtUID")
Else
	u = Request.QueryString("uID")
End If	
	
strSql1 = "SELECT * FROM tblUserAccounts WHERE uid = " & "'" & u & "'"
Set rsUser = CreateObject("adodb.recordset")
rsUser.Open strSql1, conn, 3, 3
		


		strSql2 = "SELECT * FROM tblCPMembers WHERE uID = " & "'" & rsUser.Fields("uid") & "'"
		Set rsCP = CreateObject("adodb.recordset")
		rsCP.Open strSql2, conn, 3, 3

Response.Write "<center><table border=1 bordercolor=darkblue bgcolor=lightyellow cellpadding=1 cellspacing=1 STYLE=" & """table-layout:auto;border-collapse:collapse""" & " bgcolor=WHITE wrap=true>"
Response.Write  "<tr colspan=8><th colspan=8><center><b><h5>" & "Custom Pat Orders for " & rsCP.Fields("uID") & "</h5></b></center></th></tr>" 	
Response.Write  "<tr colspan=8><th colspan=8><center><b><h5>" & "<a href=" & """AdminLogin.asp""" & ">" & "HOME" & "</a></h5></b></center></th></tr>" 	

Response.Write "<tr bgcolor=lightgreen><th>Edit</th><th>Pattern Name</th><th>Date</th><th>Desired Width</th><th>Desired Height</th><th>Floss</th><th>ReadyYN</th><th>Payment Status</th></tr>"
	Do While Not rsCP.EOF
	If rsCP.Fields("PaymentStatus") <> "None" Then
		Response.Write "<tr bgcolor=lightblue>"
	Else
		Response.Write "<tr>"
	End If	
		Response.Write "<td><a href=" & "'" & "console_EDIT_custPat_members.asp?cpID=" & rsCP.Fields("cpID") & "'" & ">" & "<B>EDIT</B>" & "</a></td>"		
		Response.Write "<td>" & rsCP.Fields("cpName") & "</td>"
		Response.Write "<td>" & rsCP.Fields("InitDate") & "</td>"
		Response.Write "<td>" & rsCP.Fields("patWidth") & "</td>"
		Response.Write "<td>" & rsCP.Fields("patHeight") & "</td>" 
		Response.Write "<td>" & rsCP.Fields("FlossType") & "</td>" 
		Response.Write "<td>" & rsCP.Fields("ReadyYN") & "</td>"	
		
		If 	rsCP.Fields("PaymentStatus") = "Full" Then
			Response.Write "<td bgcolor=yellow>" & rsCP.Fields("PaymentStatus") & "</td>" 
		ElseIf rsCP.Fields("PaymentStatus") = "Half" Then		
			Response.Write "<td bgcolor=red>" & rsCP.Fields("PaymentStatus") & "</td>" 
		Else
			Response.Write "<td>" & rsCP.Fields("PaymentStatus") & "</td>" 		
		End If	
		Response.Write "</tr>"
	rsCP.MoveNext
	Loop
Response.Write "</table></center>"

rsCP.Close
Set rsCP = nothing

conn.Close
Set conn = nothing

'Response.Redirect("MemberServices.asp?u=" & "'" & u & "'&u=" & "'" & p & "'")
'Response.Redirect("EmailPicture_Members.asp?u=" & "'" & u & "'&u=" & "'" & p & "'")
%>

</body>
</html>











