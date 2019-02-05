<%
	'Include in every page, to ensure valid u has logged in
	
	'######################   Verify sessionid and reset expire   ###############
	'strProvider = "driver={SQL Server};server=jsc-srq-irm;database=nsas;uid=;pwd="	
	'strProvider = "driver={SQL Server};server=jsc-srq-irm;database=nsas;uid=nsas;pwd=jaugustyn"	

	Dim u, p
	Dim conn2
	Dim cookie
	Dim rst1
	Dim strQuery
	Dim strDBpath

strDBpath = Server.MapPath("\db\access.mdb")


Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"


	cookie = Request.Cookies("NSAS")("sessionid")
	
	strQuery="select * from tblSessions WHERE sessionid ='"& cookie & "'"

	Set rst1 = CreateObject("adodb.recordset")
	rst1.Open strQuery, conn2, 3, 3

	If rst1.EOF Then
		
		rst1.Close
		conn2.Close
		set rst1 = nothing
		set conn2 = nothing
'		Response.Redirect("default.htm")

	Else
		
		conn2.Execute "UPDATE tblSessions SET sessiontime = '" & FormatDateTime(Now, vbGeneralDate) & "' WHERE sessionid = '" & cookie & "'" 
		conn2.Execute "DELETE FROM tblSessions WHERE sessiontime < '" & Now() - .01 & "'"

		rst1.Close
		conn2.Close
		set rst1 = nothing
		set conn2 = nothing

	End If


%>