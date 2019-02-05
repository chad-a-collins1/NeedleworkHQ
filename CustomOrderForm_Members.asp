<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>


<html>
<head>
<title>CrossStitchConnection</title></head>
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

<Script Language="javascript1.2">
	
	
	function fctClose() {
		var UID	= document.theForm.txtUID.value
		var PASS = document.theForm.txtPassword.value
	
		window.navigate("memberservices.asp?u=" + UID + "&u=" + PASS)
	}

</Script>


</head>


<body background="paper_old.gif">
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

	
	UID = Request.QueryString("u")
		
	'strSql ="SELECT * FROM tblPictures, tblCategories WHERE tblCategories.CategoryID = tblPictures.CategoryID AND tblPictures.CategoryID = " & cID 
	
		strSql = "SELECT * FROM tblUserAccounts WHERE uid = " & "'" & UID & "'"  
		
	Set rsUser = CreateObject("adodb.recordset")
	rsUser.Open strSql, conn, 3, 3
		
			
%>
<Form Action="submitMembersCustomOrderToDB.asp" Method="Post" Name="theForm" onUnLoad="getUP()">
<Table background="yellow.jpg" height="300" width="500" border=0>
<TR valign="top">
<TH colspan=2><CENTER><H4>Cross Stitch Connection Custom Pattern Order Form</H4></CENTER><BR><BR></TH>
</TR>

<TR>
<TD colspan=2>
<Font Color="red"><H5><B><U>Step 1:</U></B></font>&nbsp;<font color="black">Please fill out the information below and click "Next" when you are finished. Note: There is a $5.99 processing fee.</H5></Font>
</TD>
</TR>

<TR>
<TD><B>First Name:</B></TD><TD align="left"><Input type="text" size="25" name="txtFName" value="<%= rsUser.Fields("Fname") %>"></TD>
</TR>
<TR>
<TD><B>Last Name:</B></TD><TD align="left"><Input type="text" size="25" name="txtLName" value="<%= rsUser.Fields("Lname") %>"></TD>
</TR>

<TR>
<TD><B>UserID:</B></TD><TD align="left"><Input type="text" size="25" name="txtUID" value="<%= rsUser.Fields("uid") %>"></TD>
</TR>
<TR>
<TD><B>Email:</B></TD><TD align="left"><Input type="text" size="25" name="txtEmail" value="<%= rsUser.Fields("Email") %>"></TD>
</TR>

<TR>
<TD><P>&nbsp;</P></TD><TD><P>&nbsp;</P></TD>
</TR>
<TR>
<TD><P>&nbsp;</P></TD><TD><P>&nbsp;</P></TD>
</TR>

<TR>
<TD><B>Desired Pattern Name:</B></TD><TD align="left"><Input type="text" size="25" name="txtCPname"></TD>
</TR>
<TR>
<TD><B>Desired Floss Type:</B></TD><TD align="left"><Input type="text" size="25" name="txtFloss"></TD>
</TR>

<TR>
<TD><B>Desired Width (Stitches):</B></TD><TD align="left"><Input type="text" size="25" name="txtPatWidth"></TD>
</TR>
<TR>
<TD><B>Desired Height (Stitches):</B></TD><TD align="left"><Input type="text" size="25" name="txtPatHeight"></TD>
</TR>

<TR>
<TD align="left" colspan=2><B>Description:</B><BR><TEXTAREA cols=60 rows=4 name="Description"></TEXTAREA></TD>
</TR>

<TR>
<TD><center><input type="Submit" value="  Next >> "></center></TD><TD><center><Input type="button" value="    Cancel    " onClick="fctClose()" ></center></TD>
</TR>

</Table>
<input type="hidden" name="txtPassword" value="<%= rsUser.Fields("pswd") %>" >
</Form>
</body>

</html>
