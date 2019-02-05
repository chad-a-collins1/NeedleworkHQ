<HTML>
<HEAD>
<TITLE>New Account</TITLE>
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">

<STYLE type=text/css>
	p {text-align:justify;font-size: 8pt;font-family: "Verdana"; }
	td {font-size: 8pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 8pt;font-family: "Verdana"; color: "black";}
	a {font-size: 8pt;font-family: "Verdana"; color: "blue";}
	ul {font-size: 7pt;font-family: "Verdana"; color: "black";}
</STYLE>


<SCRIPT Language="javascript">

function submitForms() {
	if (isLName() && isFName() && isUID() && isEmail() && isPassword() && isPasswordLen() && isValidatePassword() && isValidatePassword2() && isKeyword()) { return true }
	else {return false}
	};


function isLName() {
if (document.theForm.billTo_lastName.value == "") {
alert ("\n The Last Name field is blank. \n\n Please enter your last name.")
document.theForm.billTo_lastName.focus();
return false;
}
return true;
}

function isFName() {
if (document.theForm.billTo_firstName.value == "") {
alert ("\n The First Name field is blank. \n\n Please enter your first name.")
document.theForm.billTo_firstName.focus();
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
if (document.theForm.billTo_email.value == "") {
alert ("\n The E-Mail field is blank. \n\n Please enter your E-Mail address.")
document.theForm.theForm.billTo_email.focus();
return false;
}
if (document.theForm.billTo_email.value.indexOf ('@',0) == -1 &&
document.theForm.billTo_email.value.indexOf ('.',0) == -1) {
alert ("\n The E-Mail field requires a \"@\" and a \".\"be used. \n\nPlease re-enter your E-Mail address.")
document.theForm.billTo_email.select();
document.theForm.billTo_email.focus();
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


</SCRIPT>





</HEAD>
<BODY aLink=#b8c9a6  leftMargin=0 link=#b8c9a6 
style="FONT-FAMILY: Arial; FONT-SIZE: 8pt" text="darkblue" topMargin=0 background="aida.jpg">


<CENTER>

<CENTER>
<Table width="50%" border=0 bordercolor="darkblue"><tr><td><center>

<Table width="50%" cellpadding=0 cellspacing=0>
<tr>
<td align="center"><a href="AdminLogin.asp"><img src="/images/header.jpg" border=0 alt="cross stitch"></a></td>
</tr>
</table>

</center></td></tr></table>
</CENTER>

<CENTER>



<form action="NewMember2.asp" method="post" name="theForm" >



<table width="60%" bgcolor="lightyellow"  bordercolor="silver">
<tr colspan=2><td colspan=2>
<Font size=1 color="red"><B>Please complete the form below to set up your user profile. You will then be directed to our payment processing page. <u>The membership fee to Cross Stitch Connection is a one-time charge of $9.99</u></B></font>

</td>
</tr>
<TR>
<TD size="35">Last Name:<BR><INPUT Name="billTo_lastName"  size=30 style="HEIGHT: 22px; WIDTH: 180px"></TD>
<TD size="35">First Name:<BR><INPUT Name="billTo_firstName"  size=20 style="HEIGHT: 22px; WIDTH: 180px"></TD>
</TR>

<TR>
<TD size="35">Street Address:<BR><INPUT Name="billTo_street1"  size=30 style="HEIGHT: 22px; WIDTH: 300px"></TD>
<TD size="35">City:<BR><INPUT Name="billTo_city"  size=30 style="HEIGHT: 22px; WIDTH: 100px"></TD>
</TR>

<TR>
<TD size="35">State (US only):<BR><INPUT Name="billTo_state"  size=30 style="HEIGHT: 22px; WIDTH: 180px"></TD>
<TD size="35">ZIP/Postal Code:<BR><INPUT Name="billTo_postalCode"  size=30 style="HEIGHT: 22px; WIDTH: 80px"></TD>
</TR>

<TR>
<TD size="35"><B>DESIRED USER NAME:</B><BR><INPUT Name="txtDesiredID" size=50 style="HEIGHT: 22px; WIDTH: 180px"></TD>
<TD size="35">E-mail Address:<BR><INPUT Name="billTo_email" size=45 style="HEIGHT: 22px; WIDTH: 180px"></TD>
</TR>

<TR colspan=2>
<TD colspan=2><BR><P><b><u>Please verify that your email address is correct!</u>
  Your account information and receipt will me emailed to you.</b></P><BR></TD>
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

<tr><td colspan="2"><br><font color="red" size="1"><b>We accept electronic payments via <i>Mastercard</i> and <i>Visa</i> and <i>PayPal</i>. Our electronic payment service is facilitated by <i>Bank of America's</i> merchant account provider, <i>CyberSource utilizing 128-bit encryption Secure Socket Layer security.</i></b></font></td></tr>
<tr><td colspan="2"><img src="x-click-butcc.gif">&nbsp;&nbsp;&nbsp;<img src="BOA.gif"></td></tr>
</table>
<br>
<table>
<tr><td><center><input type="submit" name="submit" value="Continue" onClick="return submitForms()" ></center></td></tr>
</table>


</CENTER>
</FORM>
</BODY>
</HTML>
