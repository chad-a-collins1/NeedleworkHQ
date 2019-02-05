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
	td {font-size: 10pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 8pt;font-family: "Verdana"; color: "darkblue";}
	a {font-size: 8pt;font-family: "Verdana"; color: "blue";}
</STYLE>

<body>

<%


	Response.Write "<Center><Font color=" & """darkblue""" & "<H3><B>Welcome to " & rst2.Fields("sName") & "</B></H3></Font></Center>"
	Response.Write "<center><a href='http://www.pcstitch.com/PatView/Download.ASP'><Font color='green'><B><I>Download PCStitch Viewer!</I></B></Font></a></CENTER><br>"


	Response.Write 	"<Center><Table cellpadding=8 border=1 background=" & "'" & "yellow.jpg" & "'" & ">"

	i = 1

	Response.Write "<tr>"
	
		
			Response.Write "<td width=150>"
			Response.Write "<Center>"
			Response.Write "<B><H4>" & "Spring Bunnies" & "</H4></B>by<BR>" & "NeedleWorkHQ" & Space(1) & "<BR><BR>"			
			Response.Write "<a href=" & "'" & "dev/Production/" & "spring_bunnies.jpg" & "'" & "><IMG SRC='" & "dev/ConsignmentShop/" & rst1.Fields("pID") & ".jpg" & "'" & "height=" & """80""" & "width=" & """100""" & "border=0></a>"
			Response.Write "<BR>"
			Response.Write "<a href=" & "'" & "dev/Production/" & "spring_bunnies_large.jpg" & "'" & ">[enlarge]</a>"
			Response.Write "<BR><BR>"
			Response.Write "<B>Price: " & rst1.Fields("pPrice") & "</B><BR><BR>"
			Response.Write "<form action=" & """https://www.paypal.com/cgi-bin/webscr""" & "method=" & """post""" & ">"
			Response.Write "<input type=" & """hidden""" & " name=" & """cmd""" & "value=" & """_xclick""" & ">"
			Response.Write "<input type=" & """hidden""" & " name=" & """business""" & " value=" & """needleworkhq@yahoo.com""" & ">"
			Response.Write "<input type=" & """hidden""" & " name=" & """item_name""" & " value=" & """Consignment""" & ">"
			Response.Write "<input type=" & """hidden""" & " name=" & """item_number""" & " value=" & "pattern-of-week" & ">"
			Response.Write "<input type=" & """hidden""" & " name=" & """amount""" & " value=" & "2.50" & ">"
			Response.Write "<input type=" & """hidden""" & " name=" & """return""" & " value=" & "http://www.needleworkhq.com/orderViewer11.asp?p=" & rst1.Fields("pID") & ">"
			Response.Write "<input type=" & """hidden""" & " name=" & """cancel_return""" & "value=" & """http://www.needleworkhq.com/orderCanceled.htm""" & ">"
			Response.Write "<input type=" & "image src=" & """http://images.paypal.com/images/x-click-but01.gif""" & "border=" & """0""" & "name=" & """submit""" & ">"
			Response.Write "</form>"
			Response.Write "</Center>"
			Response.Write "</td>"
			Response.Write "</tr><tr>"
					
Response.Write "</tr></table>"
%>
<BR>
</body>

</html>
