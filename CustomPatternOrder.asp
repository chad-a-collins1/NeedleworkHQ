

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
	a {font-size: 8pt;font-family: "Verdana"; color: "blue";}
</STYLE>

</head>


<body background="paper_old.gif">
<CENTER>

<Form Action="submitCustomOrderToDB.asp" Method="Post" Name="theForm" onUnLoad="getUP()"><CENTER>
<Table bgcolor="lightyellow" height="350" width="700" border=0 cellpadding=5 cellspacing=5>
<TR><TD><CENTER><TABLE>
<TR valign="top">
<TH colspan=4><CENTER><H3>Cross Stitch Connection</H3><BR><H4>Custom Pattern Order Form</H4></CENTER><BR><BR></TH>
</TR>

<TR>
<TD colspan=4>
<Font Color="red"><H5><B><U>Step 1:</U></B></font>&nbsp;<font color="black">Please fill out the information below and click "Next" when you are finished.</H5></Font>
</TD>
</TR>

<TR>
<TD><B>First Name:</B></TD><TD><Input type="text" size="25" name="txtFName"></TD>
<TD><B>Last Name:</B></TD><TD><Input type="text" size="25" name="txtLName"></TD>
</TR>

<TR>
<TD><B>Address1:</B></TD><TD><Input type="text" size="25" name="txtAddress1"></TD>
<TD><B>Address2:</B></TD><TD><Input type="text" size="25" name="txtAddress2"></TD>
</TR>

<TR>
<TD><B>City:</B></TD><TD><Input type="text" size="25" name="txtCity"></TD>
<TD><B>State:</B></TD><TD><Input type="text" size="25" name="txtState"></TD>
</TR>

<TR>
<TD><B>Mail Code:</B></TD><TD><Input type="text" size="25" name="txtZIP"></TD>
<TD><B>Country:</B></TD><TD><Input type="text" size="25" name="txtCountry"></TD>
</TR>

<TR>
<TD><B>UserID:</B></TD><TD><Input type="text" size="25" name="txtUID"></TD>
<TD><B>Email:</B></TD><TD><Input type="text" size="25" name="txtEmail"></TD>
</TR>

<TR>
<TD><P>&nbsp;</P></TD><TD><P>&nbsp;</P></TD>
<TD><P>&nbsp;</P></TD><TD><P>&nbsp;</P></TD>
</TR>

<TR>
<TD><P>&nbsp;</P></TD><TD><P>&nbsp;</P></TD>
<TD><P>&nbsp;</P></TD><TD><P>&nbsp;</P></TD>
</TR>

<TR>
<TD><B>Pattern Name:</B></TD><TD><Input type="text" size="25" name="txtCPname"></TD>
<TD><B>Desired Floss Type:</B></TD><TD><Input type="text" size="25" name="txtFloss"></TD>
</TR>

<TR>
<TD><B>Desired Width (Stitches):</B></TD><TD><Input type="text" size="25" name="txtPatWidth"></TD>
<TD><B>Desired Height (Stitches):</B></TD><TD><Input type="text" size="25" name="txtPatHeight"></TD>
</TR>

<TR>
<TD><P>&nbsp;</P></TD><TD><P>&nbsp;</P></TD>
<TD><P>&nbsp;</P></TD><TD><P>&nbsp;</P></TD>
</TR>

<TR>
<TD><center><Input type="button" value="    Cancel    " onClick="fctClose()" ><Center></TD><TD><P>&nbsp;</P></TD>
<TD><P>&nbsp;</P></TD><TD><center><input type="Submit" value="  Next >> "></center></TD>
</TR>
</CENTER></TD></TR></TABLE>
</Table></CENTER>

</Form>
</body>

</html>
