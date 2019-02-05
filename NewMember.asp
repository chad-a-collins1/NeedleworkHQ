

<HTML>
<HEAD>
<TITLE>New Account</TITLE>
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<META name=VI60_DTCScriptingPlatform content="Server (ASP)">
<META name=VI60_defaultClientScript content=JavaScript>
   
<SCRIPT Language="javascript">

function submitForms() {
	if (isLName() && isFName() && isUID() && isEmail() && isPassword() && isPasswordLen() && isValidatePassword() && isValidatePassword2() && isKeyword()) { return true }
	else {return false}
	};


function isLName() {
if (document.theForm.txtLname.value == "") {
alert ("\n The Last Name field is blank. \n\n Please enter your last name.")
document.theForm.txtLname.focus();
return false;
}
return true;
}

function isFName() {
if (document.theForm.txtFname.value == "") {
alert ("\n The First Name field is blank. \n\n Please enter your first name.")
document.theForm.txtFname.focus();
return false;
}
return true;
}

function isUID() {
if (document.theForm.txtDesiredID.value == "") {
alert ("\n The Desired UserID field is blank. \n\nPlease enter your desired user ID.")
document.theForm.txtDesiredID.focus();
return false;
}
return true;
}

function isEmail() {
if (document.theForm.txtEmail.value == "") {
alert ("\n The E-Mail field is blank. \n\n Please enter your E-Mail address.")
document.theForm.txtEmail.focus();
return false;
}
if (document.theForm.txtEmail.value.indexOf ('@',0) == -1 &&
document.theForm.txtEmail.value.indexOf ('.',0) == -1) {
alert ("\n The E-Mail field requires a \"@\" and a \".\"be used. \n\nPlease re-enter your E-Mail address.")
document.theForm.txtEmail.select();
document.theForm.txtEmail.focus();
return false;
}
return true;
}

function isPassword() {
if (document.theForm.txtDesiredPswrd.value == "") {
alert ("\n The Desired Password field is blank. \n\nPlease enter your desired password code.")
document.theForm.txtDesiredPswrd.focus();
return false;
}
return true;
}

function isPasswordLen() {
var strP = document.theForm.txtDesiredPswrd.value
var strLen = strP.length
if (strLen < 5) {
alert ("\n The Desired Password field must be at least 5 characters long.")
document.theForm.txtDesiredPswrd.focus();
return false;
}
return true;
}

function isValidatePassword() {
if (document.theForm.txtValidatePswrd.value == "") {
alert ("\n The Validate Password field is blank. \n\nPlease enter your password.")
document.theForm.txtValidatePswrd.focus();
return false;
}
return true;
}

function isValidatePassword2(){
var strP = document.theForm.txtDesiredPswrd.value
var strV = document.theForm.txtValidatePswrd.value
if (strP != strV) {
alert ("\n Your password validation failed.")
document.theForm.txtValidatePswrd.focus();
return false;
}
return true;
}


function isKeyword() {
if (document.theForm.txtKeywrd.value == "") {
alert ("\n The Keyword field is blank. \n\nPlease enter a keyword.")
document.theForm.txtKeywrd.focus();
return false;
}
return true;
}

function fctClose() {
	document.close();
}

//End javascript
</SCRIPT>


</HEAD>
<BODY aLink=#b8c9a6  leftMargin=0 link=#b8c9a6 
style="FONT-FAMILY: Arial; FONT-SIZE: 8pt" text="darkblue" topMargin=0 background="paper_old.gif">


<STYLE type=text/css>
	p {text-align:justify;font-size: 8pt;font-family: "Verdana"; }
	td {font-size: 9pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 8pt;font-family: "Verdana"; color: "black";}
	a {font-size: 8pt;font-family: "Verdana"; color: "blue";}
	ul {font-size: 7pt;font-family: "Verdana"; color: "black";}
</STYLE>
<CENTER>
<Table width="50%" border=0 bordercolor="darkblue"><tr><td><center>

<Table width="50%" cellpadding=0 cellspacing=0>
<tr>
<td align="center"><a href="AdminLogin.asp"><img src="crossstitch.jpg" width="550" height="65" border=0 alt="cross stitch"></a></td>
</tr>
</table>
</center></td></tr></table>
<H5><B>Membership Enrollment Form</B></H5>
</CENTER>
<CENTER>
<FORM Name="theForm" Method="POST" Action="AddMemberToDB.asp" onSubmit="return submitForms()">
<table width="60%" bgcolor="lightyellow">
<tr colspan="2"><td colspan="2"><font color="red"><H5><b><u>Step 2:</u>&nbsp;</font>Please fill out the form below to set up your user profile.</b></H5></td></tr>
<TR>
<TD size="35">Last Name:<BR><INPUT Name="txtLname"  size=30 style="HEIGHT: 22px; WIDTH: 180px"></TD>
<TD size="35">First Name:<BR><INPUT Name="txtFname"  size=20 style="HEIGHT: 22px; WIDTH: 180px"></TD>
</TR>

<TR>
<TD><P>&nbsp;</P></TD>
<TD><P>&nbsp;</P></TD>
</TR>

<TR>
<TD size="35"><B>USER NAME:</B><BR><INPUT Name="txtDesiredID" size=50 style="HEIGHT: 22px; WIDTH: 180px"></TD>
<TD size="35">E-mail Address:<BR><INPUT Name="txtEmail" size=45 style="HEIGHT: 22px; WIDTH: 180px"></TD>
</TR>

<TR colspan="2">
<TD colspan="2"><P><B>Please use at least 5 characters for your password.</B></P><BR></TD>
</TR>

<TR>
<TD size="35">Desired Password:<BR><INPUT type="password"  Name="txtDesiredPswrd" size=50 style="HEIGHT: 22px; WIDTH: 180px" ></TD>
<TD size="35">Validate Password:<BR><INPUT type="password" Name="txtValidatePswrd" size=50 style="HEIGHT: 22px; WIDTH: 180px"></TD>
</TR>

<TR>
<TD size="35">Keyword Type:<BR>
<SELECT dataSrc="" id="keywrdType" Name="cboKeywrdType" style="HEIGHT: 22px; WIDTH: 180px"> 
<OPTION selected 
        >What is your favorite color?</OPTION>
<OPTION>What is your favorite pets name?</OPTION>
<OPTION>What city where you were born in?</OPTION>
<OPTION>What is your favorite holiday?</OPTION>
</SELECT>
</TD>
<TD size="35">Keyword:<BR><INPUT Name="txtKeywrd" size=50 style="HEIGHT: 22px; WIDTH: 180px"></TD>
</TR>
<tr colspan=2><td colspan=2>
<center><input type="submit" name="submit"></center>
</form>
</td></tr>
</table>

</CENTER>
</BODY>
</HTML>































