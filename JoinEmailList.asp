<%

' The Main Subroutine
'***************************************************************************
Sub Main ()

Dim strAction
Dim strEmail
Dim intRC
Dim strOption



strAction = Request("v")


Select Case strAction
 
   Case "5":  ' try to subscribe or unsubscribe the email address
       strEmail = Request("txtEmail")
       strOption = Request("cmdSubmit")
       
       intRC = fn_SubscribeUnsubscribe(strEmail, strOption)
       If intRC = 0 Then
          DisplaySuccess strEmail, strOption
       Else
          DisplayFailure strEmail, strOption
       End if
   
   Case Else:
      ' if the value in the query string is not correct then just redirect to the home page
      Response.Redirect "http://www.mustangheaven.com"
       
End Select


End Sub 'Main



' This Sub displays a success page
'*********************************************************************
Sub DisplaySuccess(strEmail, strOption)
%>

<HTML>

<HEAD>
<TITLE>Mustang Heaven Email List Success</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

</HEAD>

<BODY>

<TABLE cellSpacing=0 cellPadding=0 border=0 height=512 width=800>

  <TR>
  <TD align=left valign=top>
  <TABLE cellSpacing=0 cellPadding=0 border=0 height=75 width=800>
          <TR><TD valign=top align=center><img src="Pics\mh_title.gif" height=67></TR>
 </TABLE>
  </TD>
  </TR>
  
  <TR>
   <TD><HR color=blue></TD>
  </TR>
  
<TR>   
<TD align=left valign=top>       
<TABLE cellSpacing=0 cellPadding=0 border=0 height=350 width=800>
   <TR>
     <TD align=left valign=top width=165>

     </TD>
     <TD align=left valign=top>
        <TABLE cellSpacing=0 cellPadding=5 border=0 height=350 width=700 bgcolor=WHITE>
        <TR>
        <!--#6699CC -->
           <TD bgcolor=#CCCCDD valign=center align=center>
           <font size=5> Success </font><br>
          <%
            If strOption = "SUBSCRIBE" Then
          %>
              <font size=3><% = strEmail %> has been added to the Mustang Heaven Email News List.<br>Thank You!</font>    
           <%
             Else
           %>    
              <font size=3><% = strEmail %> has been removed the Mustang Heaven Email News List.<br>Thank You!</font>    
           <%
             End If
           %>    
               
           </TD>      
        </TR>
        </TABLE>
     </TD>
   </TR>
</TABLE>
</TD>
</TR>

  <TR>
   <TD><HR color=blue></TD>
  </TR>

   <TR>
  <TD colspan=3><center><FONT size=2><BR>MustangHeaven.com.</FONT></center>
  </TR>

</TABLE>


</BODY>
</HTML>

<%
End Sub   'DisplaySuccess


' This Sub displays a Failure page
'*********************************************************************
Sub DisplayFailure(strEmail, strOption)
%>

<HTML>

<HEAD>
<TITLE>Mustang Heaven Email List Failure</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

</HEAD>

<BODY>

<TABLE cellSpacing=0 cellPadding=0 border=0 height=512 width=800>

  <TR>
  <TD align=left valign=top>
  <TABLE cellSpacing=0 cellPadding=0 border=0 height=75 width=800>
          <TR><TD valign=top align=center><img src="Pics\mh_title.gif" height=67></TR>
 </TABLE>
  </TD>
  </TR>
  
  <TR>
   <TD><HR color=blue></TD>
  </TR>
  
<TR>   
<TD align=left valign=top>       
<TABLE cellSpacing=0 cellPadding=0 border=0 height=350 width=800>
   <TR>
     <TD align=left valign=top width=165>

     </TD>
     <TD align=left valign=top>
        <TABLE cellSpacing=0 cellPadding=5 border=0 height=350 width=700 bgcolor=WHITE>
        <TR>
        <!--#6699CC -->
           <TD bgcolor=#CCCCDD valign=center align=center>
           <font size=5> ERROR </font><br>
           <%
            If strOption = "SUBSCRIBE" Then
          %>
              <font size=3><% = strEmail %> has NOT been added to the Mustang Heaven Email News List.<br>Sorry!</font>    
           <%
             Else
           %>    
              <font size=3><% = strEmail %> has NOT been removed the Mustang Heaven Email News List.<br>Sorry!</font>    
           <%
             End If
           %>  
           </TD>      
        </TR>
        </TABLE>
     </TD>
   </TR>
</TABLE>
</TD>
</TR>

  <TR>
   <TD><HR color=blue></TD>
  </TR>

   <TR>
  <TD colspan=3><center><FONT size=2><BR>MustangHeaven.com.</FONT></center>
  </TR>

</TABLE>


</BODY>
</HTML>

<%
End Sub   'DisplayFailure


'Insert an Email Address into the database or Deletes an Email from the db
'**********************************************************
Function fn_SubscribeUnsubscribe(ByVal strEmail, ByVal strOption)
    
    On Error Resume Next
    
    Dim dbconTmp           'As New ADODB.Connection
    Dim strSQL          'As String
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    '**************************************************
    
    'MS Access 2000/2002
    ' virtual path -- "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/database/database.mdb") 
    ' absolute path -- "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\websites\www.yourdomain.com\database\database.mdb"
    'MS Access 97
    ' virtual path -- "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & Server.MapPath("/database/database.mdb") 
    ' absolute path -- "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=d:\websites\www.yourdomain.com\database\database.mdb"
    ' virtual path -- "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/database/database.mdb")  
    ' absolute path -- "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=d:\websites\www.yourdomain.com\database\database.mdb"

    Response.Write "Before ConnectionString " & Err.Number

    dbconTmp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" &  Server.MapPath("/db/EmailList.mdb") 
    'dbconTmp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db\EmailList.mdb"
   

    dbconTmp.Open
    
    'Response.Write "<br>After Open " & Err.Number & Err.Description
    'Response.End
    
    If strOption = "SUBSCRIBE" Then
       strSQL = "INSERT INTO EmailList (email) values('" & strEmail & "')"
    Else
       strSQL = "DELETE FROM EmailList WHERE email = '" & strEmail & "'"
    End If
     
    dbconTmp.Execute(strSQL)
    
    dbconTmp.Close
    Set dbconTmp = Nothing
    
    If Err.Number <> 0 Then
       fn_SubscribeUnsubscribe = Err.Number
       Exit Function
    End If
    
    fn_SubscribeUnsubscribe = 0
    
End Function  'fn_SubscribeUnsubscribe



'*****************************************************************
Call Main


%>
