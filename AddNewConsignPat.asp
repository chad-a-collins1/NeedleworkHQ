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

<title>Add Consignment Pattern</title>  	
</head>
<%
Dim user, pass
Dim session, remove
Dim txtLoginStatus
Dim rsConsign
Dim strQuery
Dim conn

strDBpath = Server.MapPath("\db\NWHQ.mdb")


Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

userAlias = Request.QueryString("u")


strQuery = "SELECT * FROM tblUserAccounts WHERE userAlias = " & "'" & userAlias & "'"

strQuery3 = "SELECT * FROM tblShops"

Set rsValidate = Server.CreateObject("adodb.recordset")
rsValidate.Open strQuery, conn, 3, 3


Set rsShops = Server.CreateObject("adodb.recordset")
rsShops.Open strQuery3, conn, 3, 3
%>

<BODY background="paper_old.gif">
<STYLE type=text/css>
	p {text-align:justify;font-size: 8pt;font-family: "Verdana"; }
	td {font-size: 8pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 8pt;font-family: "Verdana"; color: "black";}
	a {font-size: 8pt;font-family: "Verdana"; color: "darkblue";}
</STYLE>
<center>
<Table width="65%" background="yellow.jpg">
<TR>
<TH>
<CENTER>
<B>New Consignment Pattern Procedures</B>
</CENTER><BR><BR>
</TH>
</TR>
<TR>
<TD><font color="red"><b><u>STEP 1:</u></b></font>&nbsp;<b>If you have never added a consignment pattern for sale, please open the following document and READ CAREFULLY.</b></TD>
</TR>
<TR>
<TD>
<CENTER>
<a href="CPR1.rtf"><font color="blue"><b>Consignment Pattern Rules and Requirements</b></font></a>
</CENTER>
</TD>
</TR>
<TR><TD><P>&nbsp;</P></TD></TR>
<TR>
<TD><font color="red"><b><u>STEP 2:</u></b></font>&nbsp;<b>Please provide the required information in the blank fields provided below.</b></TD>
</TR>
<TR>
<TD><BR><center>
	<Form action="AddNewConsignPat2.asp?u=<%= Request.QueryString("u") %>" method="post" name="theForm">
	<Table bgcolor="lightgrey">
	<TR>
	<TD>Owner:</TD>
	<TD align="left"><Input type="text" name="txtOwner" value="<%= rsValidate.Fields("FName")  %>"></TD>
	</TR>
	<TR>
	<TD>Pattern Title:</TD>
	<TD align="left"><Input type="text" name="txtPName" size="30"></TD>
	</TR>
	<TR>
	<TD>Initial Price (US $):</TD>
	<TD align="left"><Input type="text" name="txtPPrice" size="10"></TD>
	</TR>
	<TR>
	<TD>Initial Shop Location:</TD>
	<TD align="left">
		<select name="txtShop" size="1">
<%
Do While Not rsShops.EOF
	Response.Write "<option value=" & rsShops.Fields("Theme") & ">" & rsShops.Fields("sName") & "</option>"
	rsShops.MoveNext
Loop
%>
		</select>
		</TD>
	</TR>

	<TR>
	<TD>Description:</TD>
	<TD></TD>
	</TR>
	<TR colspan=2>
	<TD colspan=2><TEXTAREA Name="txtDescription" cols="60" rows="3"></TEXTAREA></TD>
	</TR>
	</Table>
	<CENTER><BR><INPUT TYPE="Submit" value=" NEXT >> "</CENTER>
	</Form>
</center></TD>
</TR>
</center>
</Table>
<%

rsShops.Close
Set rsShops = Nothing

rsValidate.Close
Set rsValidate = Nothing

conn.Close
Set conn = Nothing
%>
</BODY>

</html>




























