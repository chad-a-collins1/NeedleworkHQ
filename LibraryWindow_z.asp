<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<TITLE></TITLE>
<STYLE type=text/css>
	p {text-align:justify;font-size: 10pt;font-family: "Verdana"; color: "darkblue"; }
	td {font-size: 8pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 10pt;font-family: "Verdana"; color: "black";}
	a {font-size: 8pt;font-family: "Verdana"; color: "darkblue";}
</STYLE>
</HEAD>
<!--
<BODY background="paper_old.gif">
-->
<body background="aida.jpg">
<CENTER>



<Table width="50%" border=0 bordercolor="darkblue"><tr><td><center>


<Table width="50%" cellpadding=0 cellspacing=0>
<tr>
<td align="center"><a href="AdminLogin.asp"><img src="crossstitch.jpg" border=0 alt="cross stitch"></a></td>
</tr>
</table>
</center></td></tr></table>
<br>

<Table width="510" heigh="480" cellpadding=0 border=0>

<center>
<P><B>Your <a href="NewMember1.asp">
<font color="blue">membership</font></a> to Cross Stitch Connection affords you unlimited access to our ever-growing collection of patterns. Our collection currently consists of <i>over 750 patterns</i> in the following categories and is growing every week!<br></B></p> 


<Table background="yellow.jpg" width="500" heigh="480" cellpadding=1 border=1>
<TR><TH colspan=4>Pattern Library</TH></TR>

<%
Dim conn

strDBpath = Server.MapPath("\db\NWHQ.mdb")

Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

strQuery2 = "SELECT * FROM tblCategories ORDER BY cName"

Set rs = Server.CreateObject("adodb.recordset")
rs.Open strQuery2, conn, 3, 3

Dim i
i = 1
Response.Write "<tr>"


Do While Not rs.EOF
	If (i mod 4 = 0) then 
		Response.Write "<TD>"
		Response.Write "<center>"
'		Response.Write "<a href=CategoryViewer_tour.asp?c=" & rs.Fields("CategoryID") & "><B>" & rs.Fields("cName") & "</B></a>"
		Response.Write "<a href=" & """NewMember1.asp""" & "><B>" & rs.Fields("cName") & "</B></a>"
		Response.Write "</Center"
		Response.Write "</TD>"
		Response.Write "</tr><tr>"

	Else
			
		Response.Write "<td>"
		Response.Write "<Center>"
'		Response.Write "<a href=CategoryViewer_tour.asp?c=" & rs.Fields("CategoryID") & "><B>" & rs.Fields("cName") & "</B></a>"
		Response.Write "<a href=" & """NewMember1.asp""" & "><B>" & rs.Fields("cName") & "</B></a>"
		Response.Write "</Center>"
		Response.Write "</td>"
	
	End If	
		
	i = i + 1
	
	rs.MoveNext() 
	Loop
Response.Write "</tr></table></center>"
%>
</table>
</BODY>
