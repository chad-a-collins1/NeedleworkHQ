<%

Approved = Request.QueryString("u").Item(1)
uid = Request.QueryString("u").Item(2)
pass = Request.QueryString("u").Item(3) 
pID = Request.QueryString("u").Item(4) 

%>
<%

If Approved = "1" Then

		strDBpath = Server.MapPath("\db\NWHQ.mdb")

		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

		Set oRs2 = CreateObject("ADODB.Recordset")
	'	oRs2.CursorLocation = adUseClient 
		
		sSQL = "SELECT pID, ActiveYN, PaymentReceivedYN FROM tblConsignPatterns WHERE pID = " & Request.QueryString("u").Item(4)
		oRs2.Open sSQL, oConn, 3, 3

			oRs2.Fields("PaymentReceivedYN") = CBool(1)
			oRs2.Fields("ActiveYN") = CBool(1)
			oRs2.Update 

		oRs2.Close
		Set oRs2 = Nothing

		oConn.Close 
		Set oConn = Nothing
		
		Response.Redirect("ConsignMessage_Approved.asp?u=" & uid & "&u=" & pass & "&u=" & pID)
Else
		Response.Redirect("ConsignMessage_Declined.asp?u=" & uid & "&u=" & pass)		
End If		
%>
</BODY>
</html>





