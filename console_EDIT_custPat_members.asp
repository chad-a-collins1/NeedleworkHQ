<html>
<head></head>
<body bgcolor="darkblue">
<STYLE TYPE="text/css">
	p {text-align:justify;font-size: 9pt;font-family: "Verdana"; }
	td {font-size: 8pt;font-family: "Verdana";color: "black"; }
	th {font-size: 8pt;font-family: "Verdana"; color: "#003399";}
	a {font-size: 8pt;font-family: "Verdana"; color: "black";}
</STYLE>
<font color="black">
<%
strDBpath = Server.MapPath("/db/NWHQ.mdb")

cpID = CLng(Request.QueryString("cpID"))


Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

		strSql = "SELECT * FROM tblCPmembers WHERE cpID = " & cpID
		Set RS = CreateObject("ADODB.Recordset")
		RS.Open strSql, conn, 3, 3
%>
<center>
		<form action="editCPmemberInDB.asp" method="post">
		<table width="61%" border=3 bordercolor="darkblue" cellspacing=2 cellpadding=2 STYLE="table-layout:auto;border-collapse:collapse" bgcolor="lightyellow" wrap=true>
		<tr><td>Pattern Name</td><td><Input Type="text" Name="txtName" Value="<%= RS.Fields("cpName") %>"></td></tr>
		<tr><td>Date</td><td><%= RS.Fields("InitDate") %></td></tr>
		<tr><td>Desired Width</td><td><Input Type="text" Name="txtWidth" Value="<%= RS.Fields("patWidth") %>"></td></tr>
		<tr><td>Desired Height</td><td><Input Type="text" Name="txtHeight" Value="<%= RS.Fields("patHeight") %>"></td></tr>
		<tr><td>Floss Type</td><td><Input Type="text" Name="txtFloss" Value="<%= RS.Fields("FlossType") %>"></td></tr>
		<tr><td>ReadyYN</td><td><Input Type="text" Name="txtReadyYN" Value="<%= RS.Fields("ReadyYN") %>"></td></tr>
		<tr>
		<td>Payment (None / Half / Full)</td>
		<td>
		<Input Type="Select" Name="txtPayment" Value="<%= RS.Fields("PaymentStatus") %>">
		</td>
		</tr>
		<tr><td>ActiveYN</td><td><Input Type="text" Name="txtActiveYN" Value="<%= RS.Fields("ActiveYN") %>"></td></tr>
		<tr><td>Description:</td><td><textarea colspan=2 rowspan=7><%= RS.Fields("ActiveYN") %></textarea></td></tr>
		<tr><td>cpID</td><td><Input Type="hidden" Name="cpID" Value="<%= RS.Fields("cpID") %>"></td></tr>
		<tr colspan=2><td colspan=2><center><Input Type="submit" Value="Submit"></center></td></tr>
		</table>
		</form>
</font>
</body>
</html>