<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>


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
	td {font-size: 8pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 8pt;font-family: "Verdana"; color: "darkblue";}
	a {font-size: 8pt;font-family: "Verdana"; color: "blue";}
</STYLE>

<body>

<%

	Dim conn
	Dim rst1
	Dim strSql
	Dim strDBpath
	Dim i
	Dim Count

strDBpath = Server.MapPath("/db/NWHQ.mdb")

Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

	
	pID = CInt(Request.QueryString("pID"))
		
	strSql = "SELECT * FROM tblPictures WHERE ID = " & pID 
		
	Set rst1 = CreateObject("adodb.recordset")
	rst1.Open strSql, conn, 3, 3
	
	Response.Write 	"<Center><Table cellpadding=5 background='yellow.jpg' width='500'>"
	Response.Write 	"<tr><th><font color='red'><h5><u>Option 1:</u> Print Online</h5></font></th></tr>"

	i = 1
	Count = CInt(rst1.Fields("pageCount"))
	
	If Count = 1 Then
		Response.Write "<TR><TD><center><a href=dev/Production/" & rst1.Fields("Name") & "_pat.jpg" & ">" & "<h5>Pattern Page " & i & "</h5></a></center></TD></TR>" 
	Else

		Do While i <= Count
	
		Response.Write "<TR><TD><center><a href=dev/Production/" & rst1.Fields("Name") & "_pat_p" & i & ".jpg" & ">" & "<h5>Pattern Page " & i & "</h5></a></center></TD></TR>" 
	
		i = i + 1
		Loop
	End If
Response.Write "<TR><TD><P>&nbsp;</P></TD></TR>"
Response.Write "<TR><TD><center><a href=" & "'" & "dev/Production/" & rst1.Fields("Name") & "_key.rtf" & "'" & "><h5>[key]</h5></a></center></TD></TR>"
Response.Write "</table><br>"

	If rst1.Fields("ViewerYN").Value = CBool(1) Then
		Response.Write "<table background='yellow.jpg' width='500'>"
		Response.Write "<tr><th><font color='red'><h5><B><u>Option 2:</u> Open with PCStich Viewer</b></h5></font></th></tr>"
		Response.Write "<TR><TD><center><a href=dev/Production/" & rst1.Fields("Name") & ".pat" & ">" & "Pattern" & ".PAT" & "</a></center></td></tr>"
		Response.Write "</table>"
	ElseIf rst1.Fields("pdfYN").Value = CBool(1) Then
		Response.Write "<table background='yellow.jpg' width='500'>"
		Response.Write "<tr><th><font color='red'><h5><B><u>Option 3:</u> Open with Adobe Acrobat</b></h5></font></th></tr>"
		Response.Write "<TR><TD><center><a href=dev/Production/" & rst1.Fields("Name") & ".pdf" & ">" & "Pattern" & ".PDF" & "</a></center></td></tr>"
		Response.Write "</table>"
	Else
		Response.Write "<table background='yellow.jpg' width='500'>"
		Response.Write "<tr><th><font color='red'><h5><B>PCStich Viewer is not yet available for this pattern, sorry for the inconvenience.</b></h5></font></th></tr>"
		Response.Write "</table>"
	End If		
	
rst1.Close
Set rst1 = Nothing	

conn.Close
Set conn = Nothing
%>
<BR><BR><center>
If you have any questions about printing or viewing our patterns <a href="print.rtf">click here.</a></center>
</body>
</html>
