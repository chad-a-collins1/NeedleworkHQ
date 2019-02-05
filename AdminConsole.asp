<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>

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
	window.close()
	
function submitForms() {
	if (isLName()) { return true }
	else {return false}
	};


function isLName() {
if (document.theForm.txtCategories.value == "") {
alert ("Enter a category you stupid NIGGER!")
document.theForm.txtCategories.focus();
return false;
}
return true;
}	
	
}
</script>
</HEAD>
<BODY bgcolor="darkblue">
<STYLE TYPE="text/css">
	p {text-align:justify;font-size: 9pt;font-family: "Verdana"; }
	td {font-size: 12pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 12pt;font-family: "Verdana"; color: "black";}
	a {font-size: 10pt;font-family: "Verdana"; color: "black";}
</STYLE>
<CENTER>
<%

Dim user, pass
Dim session, remove
Dim txtLoginStatus
Dim rsValidate
Dim strQuery1
Dim conn

strDBpath = Server.MapPath("\db\NWHQ.mdb")


Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

'Set the UserID and Password variables equal to the user input via the txtUserName and txtPassword text fields, respectively.
user =	Request.Form("txtUID")
pass =	Request.Form("txtPassword")


'This query string is used with the recordset 'rsValidate,' it selects a record from tblUserAccounts based on the UserID and Password that the user entered
strQuery1 = "SELECT * FROM tblAdmin WHERE UID ='"& user & "' AND Password = '"& pass &"'" 
strQuery2 = "SELECT * FROM tblCategories ORDER BY cName"
strQuery3 = "SELECT DISTINCT userAlias, pOwnerFName FROM tblConsignPatterns"
strQuery4 = "SELECT DISTINCT uID FROM tblCPmembers"

'Create the recordset Object rsValidate
Set rsValidate = Server.CreateObject("adodb.recordset")
rsValidate.Open strQuery1, conn, 3, 3

Set rsCategories = Server.CreateObject("adodb.recordset")
rsCategories.Open strQuery2, conn, 3, 3

Set rsConsign = Server.CreateObject("adodb.recordset")
rsConsign.Open strQuery3, conn, 3, 3

Set rsCustom1 = Server.CreateObject("adodb.recordset")
rsCustom1.Open strQuery4, conn, 3, 3


If Not rsValidate.BOF And Not rsValidate.EOF Then

Response.Write "<Table><tr><td>"

	Response.Write "<FORM ACTION=" & """AdminPosts.asp""" & " METHOD=" & """POST""" & " NAME=" & """theForm""" & "><BR><BR>"
	Response.Write "<TABLE bgcolor=" & """lightyellow""" & ">"
	Response.Write "<TRcolspan=2><TH><Font color=red><B>INSERT PICTURE</B></Font></TH><TH><P>&nbsp;</P></TH></TR>"
	Response.Write "<TR colspan=2><TD><B>Enter Picture Name (NO EXTENSIONS)</B><BR><INPUT TYPE=" & """txtName""" & "Name=" & """txtName""" & "Size=" & """25""" & "></TD></TR>"	
	Response.Write "<TR colspan=2><TD><P>&nbsp;</P></TD></TR>"

	Response.Write "<TR colspan=2><TD><B>PDF?</B><BR><SELECT" & " Name=" & """ViewerYN""" & "><OPTION VALUE=" & """0""" & " SELECTED>" & "no" & "</OPTION>" & "<OPTION VALUE=" & """1""" & ">" & "yes" & "</OPTION>" & "</SELECT></TD></TR>"	

	Response.Write "<TR colspan=2><TD><P>&nbsp;</P></TD></TR>"
	Response.Write "<TR><TD align=left><B>Number of Pattern pages:&nbsp;</B><Input type=" & """text""" & "name=" & """pageCount""" & "value=" & """1""" & "size=" & """1""" & "></TD><TD><P>&nbsp;</P></TD></TR>"
	Response.Write "<TR colspan=2><TD><b>Height:</b><Input type=" & """text""" & "name=" & """txtHeight""" & "size=10></TD></TR>"
	Response.Write "<TR colspan=2><TD><b>Width:</b>&nbsp;<Input type=" & """text""" & "name=" & """txtWidth""" & "size=10></TD></TR>"
	Response.Write "<TR colspan=2><TD><P>&nbsp;</P></TD></TR>"
	Response.Write "<TR><TD><B>Select a Category</B><BR><SELECT Size=1 Name=" & """txtCategories""" & ">"
	
	Do While Not rsCategories.EOF
			Response.Write "<OPTION VALUE=' " & rsCategories.Fields("CategoryID") & "'>"
			Response.Write rsCategories.Fields("cName")
			Response.Write "</OPTION>"
			rsCategories.MoveNext 
	Loop
	
 	Response.Write "</SELECT></TD><TD>&nbsp;</TD></TR>"	
 	Response.Write "<TR colspan=2><TD><P>&nbsp;</P></TD></TR>"
 	Response.Write "<TR colspan=2><TD><INPUT TYPE=" & """SUBMIT""" & "VALUE=" & """ADD PICTURE""" & ">"
 	Response.Write "</TD></TR>"
	Response.Write "</FORM>"
	Response.Write "</table>"
	
	Response.Write "<FORM ACTION=" & """AdminDeletes.asp""" & "METHOD=" & """POST""" & "NAME=" & """theForm""" & "><BR><BR>"
	Response.Write "<TABLE bgcolor=" & """lightyellow""" & ">"
	Response.Write "<TR><TH><Font color=red><B>DELETE PICTURE</B></Font></TH></TR>"
	Response.Write "<TR><TD><B>Enter Picture Name to Delete (NO EXTENSIONS)</B><BR><INPUT TYPE=" & """txtName""" & "Name=" & """txtName""" & "Size=" & """25""" & "></TD></TR>"	
	Response.Write "<TR><TD><P>&nbsp;</P></TD></TR>"
	Response.Write "<TR><TD><INPUT TYPE=" & """SUBMIT""" & "VALUE=" & """DELETE PICTURE""" & " id= & ""1 name= & ""1>"
 	Response.Write "</TD></TR></TABLE>"	
	Response.Write "</FORM>"
	Response.Write "</td>"
	
' Consignment Patterns --------------------------------------------------------------------------------------------------------------------------------->
	Response.Write "<td><form action=" & """editConsignmentClients.asp""" & "method=" & """post""" & "name=" & """theForm""" & ">"
		Response.Write "<Table bgcolor=" & """lightyellow""" & "width=" & """150""" & "height=" & """150""" & ">"	
		Response.Write "<tr><th><font color=red><b>Consignment Customers</b></font></th></tr>"
		Response.Write "<tr><td><center>"
			Response.Write "<Select name=" & """customer""" & "size=1>"
			Do While Not rsConsign.EOF
				Response.Write "<option value=" & rsConsign.Fields("userAlias") & ">" & rsConsign.Fields("pOwnerFName") & "</option>" 
				rsConsign.MoveNext
			Loop
		Response.Write "</center></td></tr>"		
		 	Response.Write "</TD></TR>"
 	Response.Write "<tr><td><center><input type=" & """submit""" & "value=" & """View Profile""" & " id= & ""1 name= & ""1></center></td></tr>"
 	Response.Write "</TABLE>"	
	
	Response.Write "</form></td></tr>"

'---------------------------------------------------------------------------------------------------------------------------------------------------------------->	
'Custom Patterns --------------------------------------------------------------------------------------------------------------------------------->
	Response.Write "<td><form action=" & """consoleCheckCustomPats_Members.asp""" & "method=" & """post""" & "name=" & """theForm""" & ">"
		Response.Write "<Table bgcolor=" & """lightyellow""" & "width=" & """150""" & "height=" & """150""" & ">"	
		Response.Write "<tr><th><font color=red><b>Consignment Customers</b></font></th></tr>"
		Response.Write "<tr><td><center>"
			Response.Write "<Select name=" & """txtUID""" & "size=1>"
			Do While Not rsCustom1.EOF
				Response.Write "<option value=" & rsCustom1.Fields("uID") & ">" & rsCustom1.Fields("uID") & "</option>" 
				rsCustom1.MoveNext
			Loop
		Response.Write "</center></td></tr>"		
		 	Response.Write "</TD></TR>"
 	Response.Write "<tr><td><center><input type=" & """submit""" & "value=" & """View Profile""" & " id= & ""1 name= & ""1></center></td></tr>"
 	Response.Write "</TABLE>"	
	
	Response.Write "</form></td></tr>"
'---------------------------------------------------------------------------------------------------------------------------------------------------------------->	
	
Response.Write "</table>"	
	
Else
	Response.Redirect("default.htm")	
End If	

rsCategories.Close
Set rsCategories = Nothing

rsCustom1.Close
Set rsCustom1 = Nothing

conn.Close
Set conn = Nothing

%>

</CENTER>
</BODY>
</HTML>






















































