<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>
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

cpID = CLng(Request.Form("cpID"))


Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

		strSql = "SELECT * FROM tblCPmembers WHERE cpID = " & cpID
		Set RS = CreateObject("ADODB.Recordset")
		RS.Open strSql, conn, 3, 3
%>

<%

uID = RS.Fields("uID")

 RS.Fields("cpName") = Request.Form("txtName")
 RS.Fields("patWidth") = Request.Form("txtWidth") 
 RS.Fields("patHeight") = Request.Form("txtHeight") 
 RS.Fields("FlossType") = Request.Form("txtFloss") 
 RS.Fields("ReadyYN") = Request.Form("txtReadyYN")
 RS.Fields("PaymentStatus") = Request.Form("txtPayment") 
 RS.Fields("ActiveYN") = Request.Form("txtActiveYN") 
 RS.Update 
 
 RS.Close
 Set RS = Nothing
 
 conn.Close
 Set conn = Nothing

Response.Redirect("consoleCheckCustomPats_Members.asp?uID=" & uID)
 %>
		

</body>
</html>