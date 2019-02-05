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
style="FONT-FAMILY: Arial; FONT-SIZE: 14pt" text="darkblue" topMargin=0 background="paper_old.gif">


<STYLE type=text/css>
	p {text-align:justify;font-size: 10pt;font-family: "Verdana"; }
	td {font-size: 9pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 8pt;font-family: "Verdana"; color: "black";}
	a {font-size: 8pt;font-family: "Verdana"; color: "blue";}
	ul {font-size: 7pt;font-family: "Verdana"; color: "black";}
</STYLE>
<CENTER>
<BR>
<H3><B>Needlework Headquarters Membership Enrollment Form</B></H3>
</CENTER>
<BR>
<CENTER>
<Table width="70%">
<TR>
<TD>
<P><Font size=3 color="red"><b>Note:</b></font>&nbsp;<Font size=3 color="darkblue"><B>Please complete and print the form below.  After printing, please mail the form, with an enclosed <font color="blue"><b>check</b></font> or <font color="blue"><b>money order</b></font> for <font color="red"><b>$9.99(USD)</b></font> to the address below. We will email your receipt upon receiving your payment.</B></P></font>
</TD>
</TR>
</Table>
</CENTER>
<BR>
<BR>
<Center>
<B>C. Collins<BR>
444 East Medical Cntr Blvd Suite 106<BR>
Webster, Texas 77598</B>
</Center>
<BR>
<CENTER>

<FORM Name="theForm" Method="POST" Action="AddMemberToDB.asp" onSubmit="return submitForms()">

<table width="60%" bgcolor="lightyellow"  bordercolor="silver">
<TR>
<TD size="35">Last Name:<BR><INPUT Name="txtLname"  size=30 style="HEIGHT: 22px; WIDTH: 250px"></TD>
<TD size="35">First Name:<BR><INPUT Name="txtFname"  size=20 style="HEIGHT: 22px; WIDTH: 250px"></TD>
</TR>

<TR>
<TD><P>&nbsp;</P></TD>
<TD><P>&nbsp;</P></TD>
</TR>

<TR>
<TD size="35"><B>USER NAME:</B><BR><INPUT Name="txtDesiredID" size=50 style="HEIGHT: 22px; WIDTH: 248px"></TD>
<TD size="35">E-mail Address:<BR><INPUT Name="txtEmail" size=45 style="HEIGHT: 22px; WIDTH: 180px"></TD>
</TR>

<TR>
<TD size="35" colspan=2><BR><P><B>Please use at least 5 characters for your password.</B></P><BR></TD>
</TR>

<TR>
<TD size="35">Desired Password:<BR><INPUT type="password"  Name="txtDesiredPswrd" size=50 style="HEIGHT: 22px; WIDTH: 246px" ></TD>
<TD size="35">Validate Password:<BR><INPUT type="password" Name="txtValidatePswrd" size=50 style="HEIGHT: 22px; WIDTH: 181px"></TD>
</TR>

<TR>
<TD size="35">Keyword Type:<BR>
<SELECT dataSrc="" id="keywrdType" Name="cboKeywrdType" style="HEIGHT: 22px; WIDTH: 248px"> 
<OPTION selected 
        >What is your favorite color?</OPTION>
<OPTION>What is your favorite pets name?</OPTION>
<OPTION>What city where you were born in?</OPTION>
<OPTION>What is your favorite holiday?</OPTION>
</SELECT>
</TD>
<TD size="35">Keyword:<BR><INPUT Name="txtKeywrd" size=50 style="HEIGHT: 22px; WIDTH: 182px"></TD>
</TR>

</table>
<CENTER>
</CENTER>
</FORM>
</CENTER>
</BODY>
</HTML>




















































