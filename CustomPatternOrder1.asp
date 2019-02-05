

<html>
<head>
<title>NSAS Schedule</title></head>
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


<SCRIPT Language="Javascript">

function submitForms() {
//	if (isName() && isPhone() && isFax() && isEmail() && isCenterName() && isContrctrName() && isUID() && isPassword() && isValidatePassword() && isKeyword()) { return true }
	if (isFName() && isLName() && isAddress1() && isCity() && isState() && isZIP() && isEmail() && isPass() && isCPname() && isPWidth() && isPHeight()) { return true }
	else {return false}
	};


function isFName() {
if (document.theForm.txtFName.value == "") {
alert ("\n The First Name field is blank. \n\n Please enter your first name.")
document.theForm.txtFName.focus();
return false;
}
return true;
}

function isLName() {
if (document.theForm.txtLName.value == "") {
alert ("\n The Last Name field is blank. \n\n Please enter your last name.")
document.theForm.txtLName.focus();
return false;
}
return true;
}

function isAddress1() {
if (document.theForm.txtAddress1.value == "") {
alert ("\n The Address1 field is blank. \n\n Please enter a value for Address1.")
document.theForm.txtAddress1.focus();
return false;
}
return true;
}

function isCity() {
if (document.theForm.txtCity.value == "") {
alert ("\n The City field is blank. \n\n Please enter your city name.")
document.theForm.txtCity.focus();
return false;
}
return true;
}

function isState() {
if (document.theForm.txtState.value == "") {
alert ("\n The State field is blank. \n\n Please enter your state name.")
document.theForm.txtState.focus();
return false;
}
return true;
}

function isZIP() {
if (document.theForm.txtZIP.value == "") {
alert ("\n The ZIP Code field is blank. \n\n Please enter your ZIP code.")
document.theForm.txtZIP.focus();
return false;
}
return true;
}

function isEmail() {
if (document.theForm.txtEmail.value == "") {
alert ("\n The Email field is blank. \n\n Please enter your email address.")
document.theForm.txtEmail.focus();
return false;
}
return true;
}

function isPass() {
if (document.theForm.txtPass.value == "") {
alert ("\n The Password field is blank. \n\n Please enter your desired password.")
document.theForm.txtPass.focus();
return false;
}
return true;
}

function isCPname() {
if (document.theForm.txtCPname.value == "") {
alert ("\n The Title field is blank. \n\n Please enter a title for your pattern.")
document.theForm.txtCPname.focus();
return false;
}
return true;
}

function isPWidth() {
if (document.theForm.txtPatWidth.value == "") {
alert ("\n The Width field is blank. \n\n Please enter a desired width for your pattern.")
document.theForm.txtPatWidth.focus();
return false;
}
return true;
}

function isPHeight() {
if (document.theForm.txtPatHeight.value == "") {
alert ("\n The Height field is blank. \n\n Please enter a desired height for your pattern.")
document.theForm.txtPatHeight.focus();
return false;
}
return true;
}

function isDescription() {
if (document.theForm.Description.value == "") {
alert ("\n The Description Field. \n\n Please enter a description of your pattern.")
document.theForm.Description.focus();
return false;
}
return true;
}




function fctClose() {
	window.navigate("default.htm")
}


</SCRIPT>



</head>


<body background="paper_old.gif">
<CENTER>

<Form Action="submitCustomOrderToDB.asp" Method="Post" Name="theForm"  onSubmit="return submitForms()"><CENTER>
<Table bgcolor="lightyellow" height="350" width="700" border=0 cellpadding=5 cellspacing=5>
<TR><TD><CENTER><TABLE>
<TR valign="top">
<TH colspan=4><CENTER><H3>NeedleWork HeadQuarters</H3><H4>Custom Pattern Order Form</H4></CENTER><BR><BR></TH>
</TR>
<TR>
<TD colspan=4>
<Font Color="red"><H5><B><U>Step 1:</U></B></font>&nbsp;<font color="black">Please fill out the information below and click "Next" when you are finished.</H5></Font>
</TD>
</TR>

<TR>
<TD><B>First Name:</B></TD><TD align="left"><Input type="text" size="25" name="txtFName"></TD>
<TD><B>Last Name:</B></TD><TD align="left"><Input type="text" size="25" name="txtLName"></TD>
</TR>

<TR>
<TD><B>Address1:</B></TD><TD align="left"><Input type="text" size="25" name="txtAddress1"></TD>
<TD><B>Address2:</B></TD><TD align="left"><Input type="text" size="25" name="txtAddress2"></TD>
</TR>

<TR>
<TD><B>City:</B></TD><TD align="left"><Input type="text" size="25" name="txtCity"></TD>
<TD><B>State:</B></TD><TD align="left"><Input type="text" size="25" name="txtState"></TD>
</TR>

<TR>
<TD><B>ZIP:</B></TD><TD align="left"><Input type="text" size="25" name="txtZIP"></TD>
<TD><B>Country:</B></TD><TD align="left"><Input type="text" size="25" name="txtCountry"></TD>
</TR>

<TR>
<TD colspan=4><font size=1 color="red">Note: Your email address will be your login ID.</font></TD>
</TR>

<TR>
<TD><B>Email:</B></TD><TD align="left"><Input type="text" size="25" name="txtEmail"></TD>
<TD><B>Password:</B></TD><TD align="left"><Input type="text" size="25" name="txtPass"></TD>
</TR>

<TR>
<TD><B>Pattern Title:</B></TD><TD align="left"><Input type="text" size="25" name="txtCPname"></TD>
<TD><B>Desired Floss Type:</B></TD><TD align="left"><Input type="text" size="25" name="txtFloss"></TD>
</TR>

<TR>
<TD><B>Desired Width (Stitches):</B></TD><TD align="left"><Input type="text" size="25" name="txtPatWidth"></TD>
<TD><B>Desired Height (Stitches):</B></TD><TD align="left"><Input type="text" size="25" name="txtPatHeight"></TD>
</TR>
</table>
<table>

<TR colspan="4">
<TD><BR><Font Color="red"><B><U>DESCRIPTION:&nbsp;</U></font> <Font color="darkblue">Please type a short description of what you want your pattern to look like in the box below:</B></Font><BR>
<CENTER><TEXTAREA rows="5" cols="80" name="Description"></TEXTAREA></CENTER></TD>
</TR>
</Table>
<table>
<TR>
<TD><center><Input type="button" value="    Cancel    " onClick="fctClose()" ><Center></TD><TD><P>&nbsp;</P></TD>
<TD><P>&nbsp;</P></TD><TD><center><input type="Submit" value="  Next >> "></center></TD>
</TR>
</CENTER></TD></TR></TABLE>
</Table></CENTER>

</Form>
</body>

</html>
