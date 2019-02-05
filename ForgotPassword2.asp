<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>

<% 'This code block writes querystring values from the url to appropriate variables

Dim uid
Dim rsUser
Dim sqlStr
Dim conn
Dim pswd


	uid = Request.Form("txtUID")
	kwrd = Request.Form("txtAnswer")
	
	sqlStr = "SELECT * FROM tblUserAccounts WHERE uid = " & "'" & uid & "'"  & " AND Keyword = '" & kwrd & "'"  'Query String that selects a record based on 'UID' and 'Keyword'

	
	strDBpath = Server.MapPath("\db\NWHQ.mdb")

	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

	Set rsUser = Server.CreateObject("adodb.recordset")						'Create recordset Obj
	rsUser.Open sqlStr, conn, 3, 3
	
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

</HEAD>
<body background="yellow.jpg" text="darkblue" link="#00ff00" vlink="#00ff00" alink="#00ff00"> 

<FORM Method="post" Name="frmKey" onClick="fctForgotPassword(<%= uid %>)">
<BR>
<% 
	If Not rsUser.BOF And Not rsUser.EOF Then
	
		pswd =  rsUser.Fields("pswd")														'Assign the Keyword answer to the value in the recordset.
	
		rsUser.Close																					'Terminate recordset Obj
		Set rsUser = Nothing
	
		conn.Close																						'Terminate connection Obj
		Set conn = Nothing
		
		Response.Write "<H2>Your Password is:</H2> " & "<B><FONT COLOR=" & """red""" &"><H1>" & pswd & "</FONT></B></H1>"	
 
	Else
		Response.Write "<CENTER><B><H2><FONT COLOR=" & """darkblue""" &">" & "Keyword did not match.  Please try again.</FONT></H2></B></CENTER>"	
	End If	
 %>
 <BR>
</FORM>
<P>&nbsp;</P>

</BODY>
</HTML>
