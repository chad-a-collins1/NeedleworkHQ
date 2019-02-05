<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>


<html>
<head>
<title>NSAS Schedule</title>
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
	a {font-size: 14pt;font-family: "Verdana"; color: "blue";}
</STYLE>

</head>

<body background="images/bckgrd.gif">
<CENTER>
<Table width="50%" border=0 bordercolor="darkblue"><tr><td><center>


<Table width="50%" cellpadding=0 cellspacing=0>
<tr>
<td align="center"><a href="AdminLogin.asp"><img src="images/header.jpg" border=0 alt="cross stitch"></a></td>
</tr>
</table>
</center></td></tr></table>

<BR>
<CENTER>
<br>
<%

	Dim conn
	Dim rst1
	Dim strSql
	Dim strDBpath
	Dim i
	Dim Count
	Dim x
	
strDBpath = Server.MapPath("/db/NWHQ.mdb")

Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

	x = (Len(Request.QueryString("z")) - 4)
	pID = CLng(Right(Request.QueryString("z"), x))
		
	strSql = "SELECT * FROM tblFreePats2 WHERE fID = " & pID 
		
	Set rst1 = CreateObject("adodb.recordset")
	rst1.Open strSql, conn, 3, 3





If rst1.Fields("pdfYN").Value = CBool(0) Then	

	Response.Write 	"<Center><Table cellpadding=5 background='yellow.jpg' width='500'>"
	Response.Write 	"<tr><th><font color='red'><h3><u>Option 1:</u> Print Online</h3></font></th></tr>"

	i = 1
	Count = CLng(rst1.Fields("pageCount"))
	
	If Count = 1 Then
		Response.Write "<TR><TD><center><a href=dev/Production/" & rst1.Fields("fName") & "_pat.jpg" & ">" & "Pattern Page " & i & "</a><BR></center></TD></TR>" 
	Else

		Do While i <= Count
	
		Response.Write "<TR><TD><center><a href=dev/Production/" & rst1.Fields("fName") & "_pat_p" & i & ".jpg" & ">" & "Pattern Page " & i & "</a><BR></center></TD></TR>" 
	
		i = i + 1
		Loop
	End If
		Response.Write "<TR><TD><P>&nbsp;</P></TD></TR>"
		Response.Write "<TR><TD><center><a href=" & "'" & "dev/Production/" & rst1.Fields("fName") & "_key.rtf" & "'" & ">[key]</a></center></TD></TR>"
		Response.Write "</table><br><br><br>"
End If




	If rst1.Fields("ViewerYN").Value = CBool(1) Then
		Response.Write "<table background='yellow.jpg' width='500'>"
		Response.Write "<tr><th><font color='red'><h5><B><u>Option 2:</u> Open with PCStich Viewer</b></h5></font></th></tr>"
		Response.Write "<TR><TD><center><a href=dev/Production/" & rst1.Fields("fName") & ".pat" & ">" & "Pattern" & ".PAT" & "</a></center></td></tr>"
		Response.Write "</table>"
	ElseIf rst1.Fields("pdfYN").Value = CBool(1) Then
		Response.Write "<table background='yellow.jpg' width='500'>"
		Response.Write "<tr><th><font color='red'><h5><B><u>Option 3:</u> Open with Adobe Acrobat</b></h5></font></th></tr>"
		Response.Write "<TR><TD><center><a href=dev/Production/" & rst1.Fields("fName") & ".pdf" & ">" & "Pattern" & ".PDF" & "</a></center></td></tr>"
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
</body>

</html>










