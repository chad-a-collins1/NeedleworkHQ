<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>


<html>
<head>
<title>Cross Stitch Connection</title></head>
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
	a {font-size: 12pt;font-family: "Verdana"; color: "blue";}
</STYLE>

<body background="paper_old.gif">
<CENTER>
<%
u = Request.QueryString("u").Item(1)
p = Request.QueryString("u").Item(2)


'Response.Redirect("MemberServices.asp?u=" & "'" & u & "'&u=" & "'" & p & "'")
'Response.Redirect("EmailPicture.asp?u=" & "'" & u & "'&u=" & "'" & p & "'")
%>
<Table background="yellow.jpg">
<TR><TH><H3>Cross Stitch Connection Custom Pattern Order Form</H3></TH></TR> 
<TR>
<TR><TD><P>&nbsp;</P></TD></TR>
<TR><TD><P>&nbsp;</P></TD></TR>
<TR><TD><Font Color="red"><H3><U>STEP 2:</U></Font><Font color="black">&nbsp;If your custom pattern includes a picture, please email the picture to the following email address:</Font><font color="blue"><a href="mailto:support@crossstitchconnection">&nbsp;support@CrossStitchConnection</a></font>
<BR><BR><H3><font color="black">If you do not have an electronic copy of your picture, please mail a copy of your picture to the mailing addres below:</H3></font>
<CENTER><H4>
Cross Stitch Connection<BR>
444 East Medical Cntr Blvd Suite 106<BR>
Webster, TX 77598</H4><BR>
</CENTER>
</TD></TR>
<TR><TD><Font Color="red"><H3><U>STEP 3:</U></Font><Font color="black">&nbsp;Please click the secure PayPal button below,  the amount of $5.99 will be do at this time. </Font></TD></TR>
<TR><TD><CENTER>
<form action="https://www.paypal.com/cgi-bin/webscr" method="post">
<input type="hidden" name="cmd" value="_xclick">
<input type="hidden" name="business" value="needleworkhq@yahoo.com">
<input type="hidden" name="item_name" value="Custom Pattern Service">
<input type="hidden" name="item_number" value="2">
<input type="hidden" name="amount" value="5.99">
<input type="hidden" name="return" value="http://www.crossstitchconnection.com/orderComplete_Members.htm">
<input type="hidden" name="cancel_return" value="http://www.crossstitchconnection.com/orderCanceled.htm">
<input type="image" src="x-click-butcc.gif" border="0" name="submit" alt="Make payments with PayPal - it's fast, free and secure!">
</CENTER></TD></TR>
</Table>
<input type="hidden" name="txtUID" value="<%= u  %>">
<input type="hidden" name="txtPassword" value="<%= p  %>">
</form>
</body>
</html>











