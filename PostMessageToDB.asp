<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>
<html>
<head></head>
<body background="yellow.jpg">
<%          			

		Dim conn	                       			  	                			
		Dim strDBPath           
		Dim strSql
		
		strSql = "SELECT * FROM tblMessageBoard"          	        		

		
		strDBpath = Server.MapPath("\db\NWHQ.mdb")


		Set conn = Server.CreateObject("ADODB.Connection")
		conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

	
		Set rsMessage = Server.CreateObject("ADODB.Recordset")
   		rsMessage.Open strSql, conn, 3, 3
		
			rsMessage.AddNew
			rsMessage.Fields("Author") = Request.Form("txtAuthor") 
			rsMessage.Fields("Subject") = Request.Form("txtSubject")
			rsMessage.Fields("Message") = Request.Form("txtMessage") 
			rsMessage.Fields("PostDate") = CDate(Now())
			
			rsMessage.Update 
			
		rsMessage.Close 
		Set rsMessage = Nothing
		
		conn.Close
		Set conn = Nothing
		
    		Response.Redirect("Bottom1.asp")
%>
</body>
</html>