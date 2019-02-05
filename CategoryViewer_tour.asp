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
	td {font-size: 7pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 7pt;font-family: "Verdana"; color: "darkblue";}
	a {font-size: 7pt;font-family: "Verdana"; color: "blue";}
</STYLE>

<body bgcolor="white">
<P><H3><font color="darkblue">As a <a href="NewMember1.asp"><font color="blue" size="2"> member</font></a> of Needlework Headquarters, you will have easy access to our vast library of over 700 patterns. Our library is growing every week!</H3></font></P>
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

	Response.Write "<Center><Font color=" & """darkblue""" & "<H2><B>" & rst2.Fields("cName") & "</B></H2></Font></Center>"

	Response.Write 	"<Center><Table cellpadding=8 background=yellow.jpg border=2>"

	i = 1

	Response.Write "<tr>"
	
	Do While Not rst1.EOF

		If (i mod 4 = 0)  then 
		
			Response.Write "<td>"
			Response.Write "<Center>"
			Response.Write "<a href=" & """NewMember1.asp""" & "><IMG SRC='" & "dev/Production/" & rst1.Fields("Name") & ".jpg" & "'" & "height=" & """70""" & "width=" & """70""" & "></a>"
			Response.Write "<BR>"
			Response.Write "<a href=" & """NewMember1.asp""" & ">[enlarge]</a>"
			Response.Write "<a href=" & """NewMember1.asp""" & ">[pattern]</a>"
			Response.Write "<a href=" & """NewMember1.asp""" & ">[key]</a>"
			Response.Write "<br><br>"
			Response.Write "<b><font color=darkblue size=1>" & rst1.Fields("pWidth") & Space(1) & "(W)" & Space(1) & "X" & Space(1) & rst1.Fields("pHeight") & Space(1) & "(H)" & Space(1) & "Stitches" & "</font></b>"
			Response.Write "</Center>"
			Response.Write "</td>"
			Response.Write "</tr><tr>"
					
		Else
			
			Response.Write "<td>"
			Response.Write "<Center>"
			Response.Write "<a href=" & """NewMember1.asp""" & "><IMG SRC='" & "dev/Production/" & rst1.Fields("Name") & ".jpg" & "'" & "height=" & """70""" & "width=" & """70""" & "></a>"
			Response.Write "<BR>"
			Response.Write "<a href=" & """NewMember1.asp""" & ">[enlarge]</a>"
			Response.Write "<a href=" & """NewMember1.asp""" & ">[pattern]</a>"
			Response.Write "<a href=" & """NewMember1.asp""" & ">[key]</a>"
			Response.Write "<br><br>"
			Response.Write "<b><font color=darkblue size=1>" & rst1.Fields("pWidth") & Space(1) & "(W)" & Space(1) & "X" & Space(1) & rst1.Fields("pHeight") & Space(1) & "(H)" & Space(1) & "Stitches" & "</font></b>"
			Response.Write "</Center>"
			Response.Write "</td>"
	
		End If	
		
	i = i + 1
	
	rst1.MoveNext() 
	Loop

Response.Write "</tr></table>"
%>

</body>

</html>
