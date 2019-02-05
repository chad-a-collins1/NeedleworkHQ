<% @LANGUAGE = "VBScript"%>
<% Response.Buffer = True %>

<html>
<head>
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<META name=VI60_DTCScriptingPlatform content="Server (ASP)">
<META name=VI60_defaultClientScript content=JavaScript>

<title>Member Services</title>  	


</head>
<%

Dim user, pass
Dim session, remove
Dim txtLoginStatus
Dim rsValidate
Dim strQuery1
Dim conn

strDBpath = Server.MapPath("\db\NWHQ.mdb")


Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

'Set the UserID and Password variables equal to the user input via the txtUserName and txtPassword text fields, respectively.
user = Request.Form("txtUID")
pass = Request.Form("txtPassword")

If Request.QueryString("u").Count = 0 Then
	strQuery1 = "SELECT * FROM tblUserAccounts WHERE uid ='" & user & "' AND pswd = '"& pass &"'" & " AND ApprovedYN = " & "Yes"
ElseIf Request.QueryString("u").Count = 2 Then
	user = Request.QueryString("u").Item(1)
	pass = Request.QueryString("u").Item(2)
	strQuery1 = "SELECT * FROM tblUserAccounts WHERE uid = '" & user & "' AND pswd = '" & Request.QueryString("u").Item(2) & "' AND ApprovedYN = Yes"  
ElseIf Request.QueryString("u").Count = 1 Then
	strQuery1 = "SELECT * FROM tblUserAccounts WHERE uid =" & user & " AND pswd = " & pass & " AND ApprovedYN = Yes"
End If

'This query string is used with the recordset 'rsValidate,' it selects a record from tblUserAccounts based on the UserID and Password that the user entered
strQuery2 = "SELECT * FROM tblCategories ORDER BY cName"

If Request.QueryString("u").Count = 2 Then
	strQuery3 = "SELECT * FROM tblCPmembers WHERE UID = " & "'" & user & "'"
ElseIf Request.QueryString("u").Count = 1 Then
	strQuery3 = "SELECT * FROM tblCPmembers WHERE UID = " & user 	
ElseIf Request.QueryString("u").Count = 0 Then
	strQuery3 = "SELECT * FROM tblCPmembers WHERE UID = " & "'" & user & "'"
End If

'Create the recordset Object rsValidate
Set rsValidate = Server.CreateObject("adodb.recordset")
rsValidate.Open strQuery1, conn, 3, 3

Set rs = Server.CreateObject("adodb.recordset")
rs.Open strQuery2, conn, 3, 3

Set rsCustom = Server.CreateObject("adodb.recordset")
rsCustom.Open strQuery3, conn, 3, 3


'If the username and password entered by the user are valid this decision block creates a session string by concatenating 16 random 
'characters consisting of numbers, upper and lower case letters that are generated using a random number generator that randomly 
'selects a number from 1-9, and any upper case letter and lower case letter in the alphabet every execution throught the loop and then
'with three random characters of each type that are selected per loop iteration, one of them is selected randomly to be added to the 16 
'character cookie string.
If Not rsValidate.EOF Then
	session = ""
	Randomize
	For i = 1 to 16
	  intNum = Int(10 * Rnd + 48)
	  intUpper = Int(26 * Rnd + 65)
	  intLower = Int(26 * Rnd + 97)
	  intRand = Int(3 * Rnd + 1)
	  Select Case intRand
	    Case 1
	      strPartPass = Chr(intNum)
	    Case 2
	      strPartPass = Chr(intUpper)
	    Case 3
	      strPartPass = Chr(intLower)
	    End Select
	  session = session & strPartPass
	Next
	
	'Change the PageID in the QueryString that will be created on re-submittal			
	pid = "FromLogin"
	
	
	'conn.Execute "DELETE FROM tblSessions WHERE sessiontime < '" & Now() - .01 & "'" '& " OR uid = " & "'" & user & "'"
	
	'conn.Execute "INSERT INTO tblSessions (sessionid, sessiontime, uid) VALUES ('"& session &"', '"& Now() &"', '" & user & "')" 
	
	'Response.Cookies("sessionid") = session
	
	fName = rsValidate.Fields("fName")
	userAlias = rsValidate.Fields("userAlias")
	
	strQuery9 = "SELECT * FROM tblConsignPatterns WHERE userAlias = " & "'" & rsValidate.Fields("userAlias") & "'"

	Set rsMoney = Server.CreateObject("adodb.recordset")
	rsMoney.Open strQuery9, conn, 3, 3
	

'Else, If the PageID is equal to "FromHome", the login status is "OPEN" and the Form defintion tag is written so that the action taken upon submission via
'the "OK" button, will direct the user back to this page with the Values the user entered in the two text fields posted to the form so that this decision
'block may evaluate the user input and respond accordingly. The submission of this form below will change the PageID to "FromLogin"

ElseIf rsValidate.EOF Then
	Response.Redirect("LoginFailed.asp?")
		
	rsValidate.Close
	Set rsValidate = Nothing
	
End If
'session.Abandon 
%>
<body text="darkblue" link="blue" vlink="blue" alink="blue" style="FONT-FAMILY: Arial; FONT-SIZE: 10pt"  topMargin=0 background="paper_old.gif"> 
<STYLE type=text/css>
	p {text-align:justify;font-size: 10pt;font-family: "Verdana"; }
	td {font-size: 7pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 7pt;font-family: "Verdana"; color: "black";}
	a {font-size: 7pt;font-family: "Verdana"; color: "darkblue";}
</STYLE>
<CENTER>
<Table width="30%" border=0 bordercolor="darkblue"><tr><td><center>

<Table width="50%" cellpadding=0 cellspacing=0>
<tr>
<td align="center"><a href="AdminLogin.asp"><img src="crossstitch.jpg" border=0 alt="cross stitch"></a></td>
</tr>
</table>
</center></td></tr></table>

<CENTER>
<% Response.write "<BR><CENTER><H5>" & "Welcome to Member Services " & fName & "!" & "</H5></CENTER>" %>
<CENTER>
<a href="default.html"><B>Logout</B></a>
<Table cellpadding=4 cellspacing=4 height="200" >
<TR valign="top">
<TD>
<Table background="yellow.jpg" width="360" height="200" cellpadding=1 border=1>
<TR><TH colspan=4>Pattern Library</TH></TR>

<%
Dim i
i = 1
Response.Write "<tr>"


Do While Not rs.EOF
	If (i mod 4 = 0) then 
		Response.Write "<TD>"
		Response.Write "<center>"
		Response.Write "<a href=CategoryViewer.asp?c=" & rs.Fields("CategoryID") & "><B>" & rs.Fields("cName") & "</B></a>"
		Response.Write "</Center"
		Response.Write "</TD>"
		Response.Write "</tr><tr>"

	Else
			
		Response.Write "<td>"
		Response.Write "<Center>"
		Response.Write "<a href=CategoryViewer.asp?c=" & rs.Fields("CategoryID") & "><B>" & rs.Fields("cName") & "</B></a>"
		Response.Write "</Center>"
		Response.Write "</td>"
	
	End If	
		
	i = i + 1
	
	rs.MoveNext() 
	Loop
Response.Write "</tr></table>"
%>
<!-- <TD>
<Table background="aida.jpg" height=200 width=180>
<TR><TD></TD></TR>
</Table>
</TD> -->

<TD>
<Table background="yellow.jpg" height=100 width=180 border=1>
<TR>
<TD>
<FORM Action="viewCustomPatternOrder.asp" Method="post" Name="cpViewer">

<TABLE background="yellow.jpg" height=100 width=180 border=0>
<TR valign=top>
<TD>
<Center>
<H6><font color="black"><B><%= fName & "'s " %> Custom Patterns</B></font></H6>
<H6><B>Patterns on File:</B></H6>
<SELECT Size=1 Name="cpOnFile" style="WIDTH: 160px">
<%	
Dim j
Dim cost

j = 0

	Do While Not rsCustom.EOF
			Response.Write "<OPTION VALUE=' " & rsCustom.Fields("cpID") & "'>"
			Response.Write rsCustom.Fields("cpName")
			Response.Write "</OPTION>"
			If rsCustom.Fields("ReadyYN") = 1 Then
				j = j + 1
			End If
			rsCustom.MoveNext 
	Loop
	
	cost = j * 5
	
	rsCustom.Close
	Set rsCustom = Nothing
	

%>	
</SELECT>
<br>
<input type="submit" value="View" name="view">
<br><br><br><br>

</Center>

<CENTER><a href="http://www.pcstitch.com/PatView/Download.ASP"><Font color="green"><B><I>Download PCStitch Viewer!</I></B></Font></a></CENTER>
<BR>
<CENTER><a href="CustomOrderForm_Members.asp?u=<%= user %>"><Font color="blue"><B>Order a Custom Pattern</B></Font></a></CENTER>
</TD>
</TR>
</TABLE>

</Form>
</TD>
</TR>
<TR>
<TD bgcolor="lightblue">
<center>
<TABLE bgcolor="lightblue" border=0 height="7">
<TR><TD valign="top"><Center><a href="Message_Board_Frames.asp?UID=<% Response.Write user & "&UID=" & pass %>"><H5><font color="red"><B>Message Board</B></font></H5></a></Center></TD></TR>
</TABLE>
</center>
</TD>
</TR>
</Table>
</TD>
<TD valign=top>
<Table background="yellow.jpg" height="100" width="180" border=1>
<tr>
<td valign=top>
<CENTER><H6><font color="black">Consignment Account</font></H6></CENTER>
<BR>
<CENTER>
	<table border=0>
	<%
	Dim Total
	Dim TotalSales
	
		Total = 0.00
		TotalSales = 0
		
		Do While Not rsMoney.EOF 

				strQuery8 = "SELECT * FROM tblConsignSales WHERE pID = " & rsMoney.Fields("pID")
				
				Set rsSales = Server.CreateObject("adodb.recordset")
				rsSales.Open strQuery8, conn, 3, 3
				
				Dim k
				Dim Views
				
				k = 0
					Do While Not rsSales.EOF
						k = k + 1
						rsSales.MoveNext 
					Loop
					
				rsSales.Close
				Set rsSales = Nothing
								
				TotalSales = (CInt(TotalSales) + CInt(k))
				
				Dim s
				Dim Revenue
				s = cCur(rsMoney.Fields("pPrice"))
				
				Revenue = s * k

				Total = CCur(Total + Revenue)
				
				Fee = CCur(Total * 0.15)
				
				Net = CCur(Total - Fee)
			
				Views = CInt(Views) + CInt(rsMoney.Fields("pViews"))	
				
			rsMoney.MoveNext 
			row = row + 1
		Loop

	'rsMoney.Close 
	'Set rsMoney = Nothing
	

	%>
	<tr><th colspan=2><center><font color="darkblue"><H6>click the links for details</H6></font></center></th></tr>
	<tr><td><font color="blue">Total Views:</font></a></td><td><font color="blue">
	<B><%
	If CInt(Views) = 0 Then
		Response.Write "0"
	Else
		Response.Write Views
	End If
	 %></B></font></td></tr>
	<tr><td><font color="blue">Total Sales:</font></a></td><td><font color="blue"><B>
	<%
	If CCur(TotalSales) = 0.0 Then
		Response.Write "N/A"
	Else
		Response.Write TotalSales
	End If 
	%></B></font></td></tr>
	<tr><td><font color="blue">Gross Revenue:</font></a></td><td><font color="blue"><B>	
	<%
	If CCur(Total) = 0.0 Then
		Response.Write "N/A"
	Else
		Response.Write "$ " & Total
	End If 
	 %></B></font></td></tr>
	<tr><td><font color="blue">Processing Fee:</font></a></td><td><font color="blue"><B>
	<%
	 If CCur(Total) = 0.0 Then
		Response.Write "N/A"
	Else
		Response.Write "-$ " & Fee
	End If 
	%></B></font></td></tr>
	<tr><td><a href="AccountDetails.asp?g=<%= userAlias %>"><font color="red"><B>Net Revenue:</B></font></a></td>
	<td><a href="AccountDetails.asp?g=<%= userAlias %>"><B><font color="red"><U>
	<%
	 If CCur(Total) = 0.0 Then
		Response.Write "N/A"
	Else
		Response.Write "$ " & Net
	End If 
	%></U></font></B></a></td></tr>
	</table>
<BR>
<a href="AccountDetails.asp?g=<%= userAlias %>"><b><i><font color="green" size="4">$</font><font color="blue">Check Your Revenue</font><font color="green" size="4">$</font></i></B></a>
<BR>
<a href="ManageAccount.asp?g=<%= userAlias %>"><font color="blue"><B><H5>Consignment Account</H5></B></Font></a>
</CENTER>
</td>
</tr>
</Table>
</TD>


</CENTER>
</TR>
</Table>
<%
	'Close the record set rsValidate
	rsValidate.Close
	Set rsValidate = Nothing
	
	conn.Close
	Set conn = Nothing
%>
</body>
</html>
































