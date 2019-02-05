<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>


<html>
<head>
<title>Free Cross Stitch Patterns</title>
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">


<STYLE TYPE="text/css">
	p {text-align:justify;font-size: 9pt;font-family: "Verdana"; }
	td {font-size: 8pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 8pt;font-family: "Verdana"; color: "darkblue";}
	a {font-size: 8pt;font-family: "Verdana"; color: "blue";}
</STYLE>

</head>

<body background="paper_old.gif">
<CENTER>
<H3><B><FONT COLOR="darkred">Free Cross Stitch Patterns by Cross Stitch Connection</FONT></B></H3>
</CENTER>
<!--
<Table width="50%" border=0 bordercolor="darkblue"><tr><td><center>
<Table width="50%" cellpadding=0 cellspacing=0>
<tr>
<td align="center"><a href="AdminLogin.asp"><img src="crossstitch.jpg" width="550" height="65" border=0 alt="cross stitch"></a></td>
</tr>
</table>
</center></td></tr></table>
-->

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
		
	strSql = "SELECT * FROM tblFreePats2"
		
	Set rst1 = CreateObject("adodb.recordset")
	rst1.Open strSql, conn, 3, 3
	
	Response.Write 	"<Center><Table cellpadding=18 background=" & """yellow.jpg""" & "border=2>"

	i = 1

	Response.Write "<tr>"
	
	Do While Not rst1.EOF
	  If i > 16 then
		Exit Do
	  Else	
		If (i mod 4 = 0)  then 
		
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

rst1.Close
Set rst1 = Nothing

conn.Close
Set conn = Nothing
%>
</Font>

<H5><a href="NewMember1.asp">[ More Patterns ]</a></H5>
<BR><BR><BR>

</CENTER>
</body>

</html>

















