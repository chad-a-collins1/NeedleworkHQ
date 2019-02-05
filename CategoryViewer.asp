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

<body background="aida.jpg">

<%

	Dim conn
	Dim rst1
	Dim strSql
	Dim strDBpath
	Dim i

strDBpath = Server.MapPath("/db/NWHQ.mdb")

Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

	
	cID = CInt(Request.QueryString("c"))
		
	'strSql ="SELECT * FROM tblPictures, tblCategories WHERE tblCategories.CategoryID = tblPictures.CategoryID AND tblPictures.CategoryID = " & cID 
	
		strSql = "SELECT * FROM tblPictures WHERE CategoryID = " & cID 
		strSql2 = "SELECT * FROM tblCategories WHERE CategoryID = " & cID
		
	Set rst1 = CreateObject("adodb.recordset")
	rst1.Open strSql, conn, 3, 3
	
	Set rst2 = CreateObject("adodb.recordset")
	rst2.Open strSql2, conn, 3, 3

	Response.Write "<Center><Font color=" & """darkblue""" & "<H3><B>" & rst2.Fields("cName") & " Patterns</B></H3></Font></Center>"

	Response.Write "<Center><Table cellpadding=10 border=1 bgcolor=white width=" & """90%""" & ">"

	i = 1

	Response.Write "<tr>"
	
	Do While Not rst1.EOF

		If (i mod 4 = 0)  then 
		
			Response.Write "<td background=yellow.jpg>"
			Response.Write "<Center>"
'			Response.Write "<b><font color=darkred size=2>" & rst1.Fields("Name") & "</font></b><br><br>"
			Response.Write "<a href=ItemViewer.asp?pID=" & rst1.Fields("ID") & "><IMG SRC='" & "dev/Production/" & rst1.Fields("Name") & ".jpg" & "'" & "height=" & """75""" & "width=" & """75""" & "></a>"
			Response.Write "<BR>"
			Response.Write "<a href=" & "'" & "dev/Production/" & rst1.Fields("Name") & ".jpg" & "'" & ">[enlarge]</a>"
			Response.Write "<a href=" & "'" & "ItemViewer.asp?pID=" & rst1.Fields("ID") & "'" & ">[pattern]</a>"
			Response.Write "<a href=" & "'" & "dev/Production/" & rst1.Fields("Name") & "_key.rtf" & "'" & ">[key]</a>"
			Response.Write "<br><br>"
			Response.Write "<b><font color=darkblue size=1>" & rst1.Fields("pWidth") & Space(1) & "(W)" & Space(1) & "X" & Space(1) & rst1.Fields("pHeight") & Space(1) & "(H)" & Space(1) & "Stitches" & "</font></b>"
			Response.Write "</Center>"
			Response.Write "</td>"
			Response.Write "</tr><tr>"
					
		Else
			
			Response.Write "<td background=yellow.jpg>"
			Response.Write "<Center>"
'			Response.Write "<b><font color=darkred size=2>" & rst1.Fields("Name") & "</font></b><br><br>"			
			Response.Write "<a href=ItemViewer.asp?pID=" & rst1.Fields("ID") & "><IMG SRC='" & "dev/Production/" & rst1.Fields("Name") & ".jpg" & "'" & "height=" & """75""" & "width=" & """75""" & "></a>"
			Response.Write "<BR>"
			Response.Write "<a href=" & "'" & "dev/Production/" & rst1.Fields("Name") & ".jpg" & "'" & ">[enlarge]</a>"
			Response.Write "<a href=" & "'" & "ItemViewer.asp?pID=" & rst1.Fields("ID") & "'" & ">[pattern]</a>"
			Response.Write "<a href=" & "'" & "dev/Production/" & rst1.Fields("Name") & "_key.rtf" & "'" & ">[key]</a>"
			Response.Write "<br><br>"
			Response.Write "<b><font color=darkblue size=1>" & rst1.Fields("pWidth") & Space(1) & "(W)" & Space(1) & "X" & Space(1) & rst1.Fields("pHeight") & Space(1) & "(H)"  & Space(1) & "Stitches" & "</font></b>"		
			Response.Write "</Center>"
			Response.Write "</td>"
	
		End If	
		
	i = i + 1
	
	rst1.MoveNext() 
	Loop
Response.Write "</tr></table>"
%>
<BR><BR>
<H5>If you have any questions about printing or viewing our patterns <a href="print.rtf">click here.</H5></a>

</body>

</html>
