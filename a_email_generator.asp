<html>
<head>
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<META name=VI60_DTCScriptingPlatform content="Server (ASP)">
<META name=VI60_defaultClientScript content=JavaScript>
</head>
<body bgcolor="lightyellow">
<STYLE TYPE="text/css">
	p {text-align:justify;font-size: 9pt;font-family: "Verdana"; }
	td {font-size: 10pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 12pt;font-family: "Verdana"; color: "darkblue";}
	a {font-size: 8pt;font-family: "Verdana"; color: "blue";}
</STYLE>

<center>
<table cellpadding=1 cellspacing=1 border=1>
<tr bgcolor="lightblue"><th>#</th><th>Email Address</th><th>User ID</th><th>First Name</th><th>Last Name</th><th>Completed</th><tr>
<%
j = 0
row = 1
Do While j < 1001

	Randomize
	For i = 1 to 16
	  intNum = Int(10 * Rnd + (48))
	  intUpper = Int(26 * Rnd + (65))
	  intLower = Int(26 * Rnd + (97))
	  intRand = Int(3 * Rnd + 1)
	  Select Case intRand
	    Case 1
	      strPartPass = Chr(intNum)
	    Case 2
	      strPartPass = Chr(intUpper)
	    Case 3
	      strPartPass = Chr(intLower)
	    End Select
	  z_Name = z_Name & strPartPass
	Next

	If row mod 2 = 0 Then
				
		Response.Write "<tr bgcolor=white>"
			
	Else
				
		Response.Write "<tr bgcolor=lightgreen>"
				
	End If
			
Response.Write "<td>" & j & "</td><td>" & Left(z_Name,5) & "_" & j & "@yahoo.com" & "</td><td><b>" & Left(z_Name,5) & "_" & j & "</b></td><td><b>" & Left(z_Name, 5) & "</b></td><td><b>" & Right(z_Name, 7) & "</b></td>" & "<td><input type=checkbox></td></tr>" 
	z_Name=""
	row = row + 1
	j = j + 1
	i=1
Loop
%>
</table>
</center>

'	strDBpath = Server.MapPath("\db\NWHQ.mdb")
'	Set conn = Server.CreateObject("ADODB.Connection")
'	conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

'	Set rsNewAccount = Server.CreateObject("ADODB.Recordset")
'   	rs.Open strSql, conn, 3, 3
		

'	rs.AddNew 
'       rs.Update
'       rs.MoveFirst 	       
        
'        rs.Close
'        Set rs = Nothing
        
'        conn.Close
'        Set conn = Nothing
		


</body>
</html>