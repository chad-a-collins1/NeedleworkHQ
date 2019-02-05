<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>



<%

'Const GLOB_EMAILHOST = "mail.bayareaconsulting.biz"
'Const GLOB_DEFAULT_EMAILNAME = "BayAreaConsulting"
'Const GLOB_DEFAULT_EMAIL = "Support@bayareaconsulting.biz"

Const GLOB_EMAILHOST = "207.150.192.13"
Const GLOB_DEFAULT_EMAILNAME = "Cross Stitch Connection"
Const GLOB_DEFAULT_EMAIL = "support@crossstitchconnection.com"
Const GLOB_ALLOW_ASPQMAIL = "yes"
Const ALLOW_ASPQMAIL = "yes"


'Function sends an email
'**************************************************************************
Function fn_SendEmail(ByVal strTo, ByVal strFrom, ByVal strSubject, ByVal strBody)
   
    
     Dim Mailer
     Dim blnMailSent
     
     Set Mailer = Server.CreateObject("CDONTS.NewMail")
     Mailer.BodyFormat = 0
     Mailer.MailFormat = 0
     Mailer.To = strTo
     Mailer.From   = strFrom
     Mailer.Subject    = strSubject
     Mailer.Body   = strBody 
     Mailer.Importance = 1
     Mailer.Send        
     
     Set Mailer = Nothing
     
End Function  'Send email

%>


<%
	Dim conn
	Dim rsNewAccount
	Dim txtLast
	Dim txtFirst          				
	Dim txtEmail
	Dim txtDesiredID
	Dim txtDesiredPswrd
	Dim txtValidatePswrd
	Dim cboKeywrdType
	Dim txtKeywrd
	Dim strSql
	Dim rsAppInfo
	Dim txtPageID
	Dim txtStreet
	Dim txtCity
	Dim txtState
	Dim txtPostal

	'Define a SQL string to select every field from the tblNewAccount in the SAS Database 
	strSql = "Select * from tblUserAccounts"

	txtLast = Request.Form("billTo_lastName")
	txtFirst = Request.Form("billTo_firstName")											
	txtEmail = Request.Form("billTo_email")										
	txtDesiredID = Request.Form("txtDesiredID")					
	txtDesiredPswrd = Request.Form("txtDesiredPswrd")		
	txtValidatePswrd = Request.Form("txtValidatePswrd")		
	cboKeywrdType =  Request.Form("cboKeywrdType")									
	txtKeywrd =  Request.Form("txtKeywrd")		
	txtStreet = Request.Form("billTo_street1")	
	txtCity = Request.Form("billTo_city")	
	txtState = Request.Form("billTo_state")	
	txtPostal = Request.Form("billTo_postalCode")	

	
	'The following sets the PageID as read from the URL
	txtPageID = Request.QueryString("PageID")
	
	'Create a connection object and define a connection string to the SAS database DSN	
	strDBpath = Server.MapPath("\db\NWHQ.mdb")

	alias = ""
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
	  alias = alias & strPartPass
	Next



	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

	Set rsNewAccount = Server.CreateObject("ADODB.Recordset")
   	rsNewAccount.Open strSql, conn, 3, 3
		
		Dim y
		y = CBool(1)
		Dim n
		n = CBool(0)

		'Add values from New Account Page text fields to corresponding tblNewAccount fields.   	    	    	    	
		rsNewAccount.AddNew 
       
		rsNewAccount("Lname") = txtLast
		rsNewAccount("Fname") = txtFirst
		rsNewAccount("Email") = txtEmail
		rsNewAccount("uid") = txtDesiredID
		rsNewAccount("pswd") = txtDesiredPswrd
		rsNewAccount("ValidatePswrd") = txtValidatePswrd
		rsNewAccount("KeywordType") = cboKeywrdType
		rsNewAccount("Keyword") = txtKeywrd
		rsNewAccount("StartDate") = CDate(Now())
		rsNewAccount("ApprovedYN") = y
		rsNewAccount("userAlias") = alias
       
        rsNewAccount.Update
        rsNewAccount.MoveFirst 	       
        
        rsNewAccount.Close
        Set rsNewAccount = Nothing
        
        conn.Close
        Set conn = Nothing

		
	   'Auto email response parameters
   		strTo = txtEmail
   		strFrom = "Cross Stitch Connection <support@crossstitchconnection.com>"
   		strSubject = "New Account Information"
   		strBody = "<p>Welcome to Cross Stitch Connection!</p>" 
   		trBody = strBody & "<p>Please keep this email and bookmark the following link for quick access to your account:</p>"
   		strBody = strBody & "<br><b>Quick Link: " & "<a href=http://www.crossstitchconnection.com/MemberServices.asp?u=" & txtDesiredID & "&u=" & txtDesiredPswrd & ">My Cross Stitch Connection Account</a></b>"
   		strBody = strBody & "<br><b>USERID: " & txtDesiredID & "</b>"
   		strBody = strBody & "<br><b>PASSWORD: " & txtDesiredPswrd & "</b>"
   		strBody = strBody & "<br><br>Happy Stitching!" 

   Call fn_SendEmail(strTo, strFrom, strSubject, strBody)		
		
		
		
	 Response.Redirect "MemberServices.asp?u=" & txtDesiredID & "&u=" & txtDesiredPswrd 
%>












