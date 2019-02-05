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
}
</script>
</HEAD>
<BODY>
<CENTER>
<%

Dim rsValidate
Dim strQuery2
Dim conn
Dim txtCategory
Dim txtName
Dim vYN

txtCategory = CInt(Request.Form("txtCategory"))
txtName = Request.Form("txtName")

strDBpath = Server.MapPath("\db\NWHQ.mdb")

Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

strQuery1 = "SELECT Name FROM tblPictures WHERE Name = " & "'" & txtName & "'"
strQuery2 = "SELECT * FROM tblPictures"

Set rsCategories = Server.CreateObject("adodb.recordset")
rsCategories.Open strQuery2, conn, 3, 3

Set rsCheck = Server.CreateObject("adodb.recordset")
rsCheck.Open strQuery1, conn, 3, 3

If Not rsCheck.BOF And Not rsCheck.EOF Then
		Response.Redirect("NameReject.asp?Name=" & "'" & txtName & "'")
		
		rsCheck.Close
		Set rsCheck = Nothing
Else
			rsCategories.AddNew 
			rsCategories.Fields("CategoryID") = CInt(Request.Form("txtCategories"))
			rsCategories.Fields("Name") = Request.Form("txtName")
			If Request.Form("ViewerYN") = "1" Then
				vYN = CBool(1)
			Else
				vYN = CBool(0)
			End If
			rsCategories.Fields("pdfYN") = vYN
			rsCategories.Fields("pageCount") = CInt(Request.Form("pageCount"))	
			rsCategories.Fields("pHeight") = CLng(Request.Form("txtHeight"))
			rsCategories.Fields("pWidth") = CLng(Request.Form("txtWidth"))	
			rsCategories.Update 
End If


rsCategories.Close
Set rsCategories = Nothing

conn.Close
Set conn = Nothing

Response.Redirect("default.htm")
%>

</CENTER>
</BODY>
</HTML>






















































