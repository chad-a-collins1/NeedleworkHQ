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
		
		strSql = "SELECT * FROM tblFreePats2 WHERE CategoryID = " & cID 
		strSql2 = "SELECT * FROM tblCategories WHERE CategoryID = " & cID
		
	Set rst1 = CreateObject("adodb.recordset")
	rst1.Open strSql, conn, 3, 3
	
	Set rst2 = CreateObject("adodb.recordset")
	rst2.Open strSql2, conn, 3, 3

'	Response.Write "<Center><Font color=" & """red""" & "<H6><B>Here are some examples of our " & rst2.Fields("cName") & "&nbsp;patterns.</B></H6></Font></Center>"

	Response.Write 	"<Center><Table cellpadding=18 background=" & """yellow.jpg""" & "border=2>"

	i = 1

	Response.Write "<tr>"
	
	Do While Not rst1.EOF
	  If i > 2 then
		Exit Do
	  Else	
		If (i mod 4 = 0)  then 
		
			Response.Write "<td>"
			Response.Write "<Center>"
			Response.Write "<br>"
			Response.Write "<font color=" & """darkred""" & "><b><h4>" & rst1.Fields("fTitle") & "</h4></b></font><br>"
			Response.Write "<a href=" & "'" & "dev/Production/" & rst1.Fields("fName") & ".jpg" & "'" & "><IMG SRC='" & "dev/Production/" & rst1.Fields("fName") & ".jpg" & "'" & "height=" & """120""" & "width=" & """160""" & "></a>"
			Response.Write "<BR>"
			Response.Write "<a href=" & "'" & "dev/Production/" & rst1.Fields("fNamef") & ".jpg" & "'" & ">[enlarge]</a>"
			Response.Write "<a href=" & "'" & "cvF2.asp?z=" & "1832" & rst1.Fields("ID") & "'" & ">[pattern]</a>"
			Response.Write "<a href=" & "'" & "dev/Production/" & rst1.Fields("fName") & "_key.rtf" & "'" & ">[key]</a>"
			Response.Write "<br><br>"
			Response.Write "<b>" & rst1.fields("fWidth") & "(W)" & Space(1) & "X" & space(1) & rst1.fields("fHeight") & "(H)"		
			Response.Write "</Center>"
			Response.Write "</td>"
			Response.Write "</tr><tr>"
					
		Else
			
			Response.Write "<td>"
			Response.Write "<Center>"
			Response.Write "<br>"
			Response.Write "<font color=" & """darkred""" & "><b><h4>" & rst1.Fields("fTitle") & "</h4></b></font><br>"
			Response.Write "<a href=" & "'" & "dev/Production/" & rst1.Fields("fName") & ".jpg" & "'" & "><IMG SRC='" & "dev/Production/" & rst1.Fields("fName") & ".jpg" & "'" & "height=" & """120""" & "width=" & """160""" & "></a>"
			Response.Write "<BR>"
			Response.Write "<a href=" & "'" & "dev/Production/" & rst1.Fields("fName") & ".jpg" & "'" & ">[enlarge]</a>"
			Response.Write "<a href=" & "'" & "cvF2.asp?z=" & "1832" & rst1.Fields("fID") & "'" & ">[pattern]</a>"
			Response.Write "<a href=" & "'" & "dev/Production/" & rst1.Fields("fName") & "_key.rtf" & "'" & ">[key]</a>"
			Response.Write "<br><br>"
			Response.Write "<b>" & rst1.fields("fWidth") & "(W)" & Space(1) & "X" & space(1) & rst1.fields("fHeight") & "(H)"		
			Response.Write "</Center>"
			Response.Write "</td>"
	
		End If	
		
	i = i + 1
	  End If	
	rst1.MoveNext() 
	Loop
Response.Write "</tr></table>"
%>
<CENTER><Font color="blue">
<%
Response.Write "<B><a href=" & """NewMember1.asp""" & ">" & "<H4>[More" & Space(1) & rst2.Fields("cName") & "]</H4></a></B>"


rst1.Close
Set rst1 = Nothing

rst2.Close
Set rst2 = Nothing

conn.Close
Set conn = Nothing
%>
</Font>

<H5>If you have any questions about printing or viewing our patterns <a href="print.rtf">click here.</H5></a>
<BR><BR><BR>

</CENTER>
</body>

</html>

















