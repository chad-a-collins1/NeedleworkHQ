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
strQuery2 = "SELECT * FROM tblConsignPatterns"

Set rsValidate = Server.CreateObject("adodb.recordset")
rsValidate.Open strQuery, conn, 3, 3


If Not rsValidate.BOF And Not rsValidate.EOF Then
Dim uid

uid = rsvalidate.Fields("uid")
pass = rsValidate.Fields("pswd")


	Set rsConsign = Server.CreateObject("adodb.recordset")
	rsConsign.Open strQuery2, conn, 3, 3	
	
	rsConsign.AddNew 
	rsConsign.Fields("pName") = Request.Form("txtPName")
	rsConsign.Fields("Location") = Request.Form("txtShop")
	rsConsign.Fields("userAlias") = userAlias
	rsConsign.Fields("ActiveYN") = CBool(0)	
	rsConsign.Fields("pPrice") = CCur(Request.Form("txtPPrice"))
	rsConsign.Fields("pInitDate") = CDate(Now())
	rsConsign.Fields("pOwnerFName") = Request.Form("txtOwner")
	rsConsign.Fields("pVotes") = CInt(0)	
	rsConsign.Fields("pViews") = CInt(0)	
	rsConsign.Update 
'	rsConsign.Close
'	Set rsConsign = Nothing	

Else
	Response.Redirect("AddNewConsignPat.asp?u=" & userAlias)
End IF	

%>

<BODY background="paper_old.gif">
<STYLE type=text/css>
	p {font-size: 8pt;font-family: "Verdana";color: "darkblue"; }
	td {font-size: 8pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 8pt;font-family: "Verdana"; color: "black";}
	a {font-size: 8pt;font-family: "Verdana"; color: "darkblue";}
</STYLE>
<BR><BR><BR><center>
<Table width="65%" background="yellow.jpg">
<TR>
<TH>
<CENTER>
<B>Upload Your JPEG Picture File for Consignment Shop Display</B>
</CENTER><BR><BR>
</TH>
</TR>
<TR>
<TD bgcolor="lightyellow"><font color="red"><b><u>STEP 3:</u></b></font>&nbsp;<b>Upload your JPEG picture file for your pattern display in the consignment shop you have selected. Use the <font color="red">"Browse"</font> feature below to find the file on your computer then click the <font color="red">"Upload JPEG"</font> button.</b><BR><BR></TD>
</TR>
<TR><TD><center>
<FORM method="post" encType="multipart/form-data" action="AddNewConsignPat3.asp?u=<%= uid & "&u=" & pass %>" id=form1 name=form1 onSubmit="return submitForms()">
	<INPUT type="File" name="File1">
	<INPUT type="Submit" value="Upload JPEG" id=Submit1 name=Submit1>
</FORM></center></TD></TR>

<TR>
<TD><BR><center>
	
</center></TD>
</TR>
</center>
</Table><BR><BR>
<%

rsValidate.Close
Set rsValidate = Nothing

conn.Close
Set conn = Nothing
%>
</BODY>

</html>

<SCRIPT Language="Javascript">

//***********This array stores the user name and JPEG entered by the user in the Login text fields below

function submitForms() {
	if (isJPEG()) { return true }
	else {return false}
	};


function isJPEG() {
if (document.form1.File1.value == "") {
alert ("\n The JPEG field is blank. \n\nPlease upload your JPEG picture file.")
document.form1.File1.focus();
return false;
}
if ((document.form1.File1.value.indexOf ('.jpg',0) == -1)) {
alert ("\n The file you attempted to upload was not a .JPG, please upload a .JPG file only.")
document.form1.File1.select();
document.form1.File1.focus();
return false;
}
return true;
}


</SCRIPT>

























