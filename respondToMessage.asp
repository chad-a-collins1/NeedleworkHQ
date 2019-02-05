<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>
<%
Dim uid
uid = Request.QueryString("UID")
%>
<HTML>
<HEAD>
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<META name=VI60_DTCScriptingPlatform content="Server (ASP)">
<META name=VI60_defaultClientScript content=JavaScript>
<TITLE></TITLE>

<script language="javascript1.2">
function fctClose() {
	window.navigate("Right.htm")
}
</script>

</HEAD>
<base target="BottomMain">
<BODY background="yellow.jpg">
<STYLE TYPE="text/css">
	p {text-align:justify;font-size: 8pt;font-family: "Verdana"; }
	td {font-size: 8pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 8pt;font-family: "Verdana"; color: "black";}
	a {font-size: 8pt;font-family: "Verdana"; color: "black";}
</STYLE>
<FORM METHOD="POST" ACTION="PostMessageToDB.asp"  NAME="theForm"> 
<CENTER>
<table border=0 width="100%">
<tr>
<td><FONT COLOR="darkblue"><B>Author:</B>&nbsp;<%= uid %></FONT></td>
</tr>
<tr>
<td><FONT COLOR="darkblue"><B>Subject:</B></FONT><BR><Input Type="Text" Name="txtSubject" size="35" value="<%= "Re: " & Request.Form("Subject") %>"></td>
</tr>
<tr>
<td>
<FONT COLOR="darkblue"><B>Message:</B></FONT><BR>
<TEXTAREA Name="txtMessage" cols="60" rows="3"></TEXTAREA>
</td>
</tr>
<tr>
<td><input type="hidden" name="txtAuthor" value="<%= uid %>"></td>
</tr>
</table>
<CENTER>
<INPUT TYPE="submit" Name="cmdSubmit" Value="   Submit   " onClick="fctClose()" >&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE="button" Name="cmdCancel" Value="   Cancel   " onClick="fctClose()">
</FORM>
</CENTER>
</CENTER>
</BODY>
</HTML>






















































