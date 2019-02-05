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

<title>Account Details</title>  	

</head>
<%
'strProvider = "driver={SQL Server};server=jsc-srq-irm;database=nsas;uid=nsas;pwd=jaugustyn"	

Dim conn
Dim userAlias
Dim strQuery
Dim strQuery2
Dim Total

userAlias = Request.QueryString("g")

strDBpath = Server.MapPath("\db\NWHQ.mdb")

Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

'This query string is used with the recordset 'rsValidate,' it selects a record from tblUserAccounts based on the UserID and Password that the user entered
strQuery = "SELECT * FROM tblConsignPatterns WHERE userAlias = " & "'" & userAlias & "'" 
Set rs = Server.CreateObject("adodb.recordset")
rs.Open strQuery, conn, 3, 3

strQuery2 = "SELECT * FROM tblUserAccounts WHERE userAlias = " & "'" & userAlias & "'" 
Set rs2 = Server.CreateObject("adodb.recordset")
rs2.Open strQuery2, conn, 3, 3

%>
<body text="darkblue" link="#00ff00" vlink="#00ff00" alink="#00ff00" style="FONT-FAMILY: Arial; FONT-SIZE: 8pt"  topMargin=0 background="paper_old.gif"> 
<STYLE type=text/css>
	p {text-align:justify;font-size: 10pt;font-family: "Verdana"; }
	td {font-size: 10pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 10pt;font-family: "Verdana"; color: "black";}
	a {font-size: 10pt;font-family: "Verdana"; color: "darkblue";}
</STYLE>
<CENTER>
<CENTER>
<Table width="50%" border=0 bordercolor="darkblue"><tr><td><center>

<Table width="50%" cellpadding=0 cellspacing=0>
<tr>
<td align="center"><a href="AdminLogin.asp"><img src="crossstitch.jpg" width="550" height="65" border=0 alt="cross stitch"></a></td>
</tr>
</table>

</center></td></tr></table><BR>

<TABLE background="yellow.jpg" border=0 bordercolor="silver" width="85%">
<TR><TD><CENTER><H3>Management Console, Account:&nbsp;<%= rs2.Fields("userAlias") %></H3></CENTER></TD></TR>
<TR valign=top>
<TD>
</CENTER><BR><BR><CENTER>
<a href="AddNewConsignPat.asp?u=<%= userAlias %>"><B><font color="blue">[Add a New Pattern to Sell]</font></B></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="MemberServices.asp?u=<%= rs2.Fields("uid") %>&u=<%= rs2.Fields("pswd") %>"><B><font color="blue">[Back to Member Services]</font></B></a><BR><BR>
<Table bgcolor="lightyellow" border=1 bordercolor="darkblue" width="80%" cellspacing=0 cellpadding=0 STYLE="table-layout:auto;border-collapse:collapse" bgcolor="WHITE" wrap=false>
<TR bgcolor="lightpink">
<% If Not rs.BOF And Not rs.EOF Then	 %>
<TH>Edit</TH>
<TH>Thumbnail</TH>
<TH>Title<INPUT TYPE="HIDDEN" VALUE="<%= rs.Fields("pID") %>"></TH>
<TH>Location</TH>
<TH>Price (USD)</TH>
<TH>Delete</TH>
</TR>
<%

		Total = 0.00
		row = 1

		Do While Not rs.EOF 
			If rs.Fields("ActiveYN") = 0 Then
				rs.MoveNext
			Else	
				If row mod 2 = 0 Then
					%>
					<tr bgcolor="lightblue">
					<%
				Else
					%>
					<tr bgcolor="lightgrey">
					<%
				End If
					%>
					<td ><center><a href="EditCSitem.asp?g=<%= rs.Fields("pID") & "&g=" & userAlias %>"><B><font color="blue">EDIT</font></B></a></center></td>	
					<td><center><B><IMG SRC="dev/ConsignmentShop/<%= rs.Fields("pID") & ".jpg" %>" height="55" width="55"></B></center></td>
					<td><center><B><%= rs.Fields("pName") %></B></a></center><Input type="hidden" name="txtID" value="<%= rs.Fields("pID")%>" ></td>
					<td ><center><B><%= rs.Fields("Location") %></B></center></td>										
					<td ><center><B>$<%= rs.Fields("pPrice") %></B></center></td>	
					<td ><center><a href="DiscontinueCSitem1.asp?u=<% Response.Write userAlias & "&u=" & rs.Fields("pID") %>"><B><font color="red">DELETE</font></B></a></center></td>																																				
				</tr>		
				<%
				rs.MoveNext 
				row = row + 1
		End If
		Loop
Else
	Response.Write "Not patterns are listed for your account"
End If

	rs.Close 
	Set rs = Nothing
	
	rs2.Close
	Set rs2 = Nothing
	
	conn.Close
	Set conn = Nothing
	
%>

<TR>
</Table>
</TD>
</TR>
<TR><TD><P>&nbsp;</P></TD></TR>
</TABLE>
</CENTER>
</body>
</html>

