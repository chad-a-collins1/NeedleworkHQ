<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>
<%
	Dim strProvider
	Dim DB
	Dim strSQL
	Dim conn
	Dim rsMessage
	Dim ID

	ID = CInt(Request.QueryString("ID").Item(1))
	UID = Request.QueryString("ID").Item(2)

	strDBpath = server.MapPath("\db\NWHQ.mdb")

	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

	Set rsMessage = Server.CreateObject("adodb.recordset")

	strSQL = "SELECT * FROM tblMessageBoard WHERE ID = " & ID

	rsMessage.Open strSQL, conn, 3, 3
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
</HEAD>
<base target="BottomMain">
<BODY background="yellow.jpg">
<CENTER>
<Form action="NewMember1.asp" method="post" name="theForm" Target="_top">
<table border=0 width="100%"><tr><td><CENTER>
<TEXTAREA Name="txtbxPurpose" cols="60" rows="6"><%= rsMessage.Fields("Message") %></TEXTAREA>
</td></tr>
<tr><td><p>&nbsp;</p></td></tr>
<tr><td><CENTER><input type="submit" value="Reply to This Message"><input type="hidden" name="Subject" value="<% = rsMessage.Fields("Subject") %>"></CENTER></td></tr>
</CENTER>
</table>
</Form>
</CENTER>
<%
rsMessage.Close
set rsMessage = Nothing
conn.Close
set conn = Nothing
%>
</BODY>
</HTML>
