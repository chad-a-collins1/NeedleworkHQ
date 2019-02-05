<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>
<% 

Dim uid
Dim user, pass
Dim rsUser
Dim sqlStr
Dim conn
Dim kwrd
Dim kwrdAnswer



		uid = Request.QueryString("UID")
	
		sqlStr = "SELECT * FROM tblUserAccounts WHERE uid = " & "'" & uid & "'"		

		strDBpath = Server.MapPath("\db\NWHQ.mdb")

		Set conn = Server.CreateObject("ADODB.Connection")
		conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"
		
%>
		
<html>
<head>
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<META name=VI60_DTCScriptingPlatform content="Server (ASP)">
<META name=VI60_defaultClientScript content=JavaScript>
<TITLE>Forgot Password</TITLE>
<SCRIPT LANGUAGE="javascript1.2">
	function fctClose() {
		window.close()
		}
</SCRIPT>
</HEAD>
<body text="darkblue" link="#00ff00" vlink="#00ff00" alink="#00ff00" background="yellow.jpg"> 

<FORM Action="ForgotPassword2.asp" Method="post" Name="frmKey">
<BR>
<BR>
<BR>
<%

		Set rsUser = Server.CreateObject("adodb.recordset")					
		rsUser.Open sqlStr, conn, 3, 3
	
	If Not rsUser.BOF And Not rsUser.EOF then 
	
		kwrd =  rsUser.Fields("KeywordType")										

		rsUser.Close																				
		Set rsUser = Nothing
	
		conn.Close																					
		Set conn = Nothing
		

%>


<% 'Create a text field and assign a value to it corresponding to the UserID passed from the login page.
	Response.Write "<B>UserID:</B>" & "<BR>" & "<INPUT Type=" & """Text""" & "width=" & """30""" & "Name=" & """txtUserName""" & "Value=" & uid & ">" & "<BR>" & "<BR>"
	Response.Write  "<B>" & kwrd & "</B><BR>" & "<BR>"
	Response.Write "<B>Type your answer below:</B>"
	Response.Write "<BR><BR><INPUT Type=" & """Text""" &  "width=" & """30""" & "Name=" & """txtAnswer""" & "><INPUT Type=" & """Submit""" & "Name=" & """Submit""" & "Value=" & """OK""" & ">"

	Else
			Response.Write "<CENTER><H2>User Name Unknown</H2><BR>"
			Response.Write "<INPUT Type=" & """Button""" & "Name=" & """Submit""" &"Value=" & """           OK           """ & "onClick=" & """fctClose()""" & "></CENTER>"
	End If
%>


<INPUT Type="Hidden" Name="txtUID" Value="<%= uid %>">
<BR>
</FORM>
<P>&nbsp;</P>


</BODY>
</HTML>
