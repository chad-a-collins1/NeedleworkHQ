<%@
LANGUAGE="VBScript"
ENABLESESSIONSTATE="True"
%>
<%
    Option Explicit
%>
  <!--#include file="adovbs.inc"-->
  <!--#include file="Constants.asp"-->
  <!--#include file="Random.asp"-->
  <!--#include file="EncryptDecrypt.asp"-->
  <!--#include file="DL_Application.asp" -->
<%

' This function returns a href based on the value of GLOB_PRODUCTION
'**********************************************************
Function fn_GetHomePageHREF()

   If GLOB_PRODUCTION = "yes" Then
      fn_GetHomePageHREF = HREF_PROD_HOME
   Else
      fn_GetHomePageHREF = HREF_TEST_HOME
   End If

End Function

' This function returns a href based on the value of GLOB_PRODUCTION
'**********************************************************
Function fn_GetConnectionString()

'Const GLOB_CS = "Provider=SQLOLEDB; Data Source=GORILLA1; Initial Catalog=PSdb; User Id=webapp; Password=admin"
'Const GLOB_CS = "Provider=SQLOLEDB; Data Source=ntsql05.propagtion.net; Initial Catalog=PSdb; User Id=passwordsupportcom; Password=passwordsupport330"
'Const GLOB_CS = "Provider=SQLOLEDB; Data Source=66.34.127.254; Initial Catalog=PSdb; User Id=passwordsupportcom; Password=passwordsupport330"

   If GLOB_PRODUCTION = "yes" Then
      fn_GetConnectionString = "Provider=SQLOLEDB; Data Source=66.34.127.254; Initial Catalog=PSdb; User Id=passwordsupportcom; Password=passwordsupport330"
   Else
      fn_GetConnectionString = "Provider=SQLOLEDB; Data Source=GORILLA1; Initial Catalog=PSdb; User Id=webapp; Password=admin"
   End If

End Function


' This function returns a href based on the value of GLOB_PRODUCTION
'**********************************************************
Function fn_GetSSLHREF(strPage)

   If GLOB_PRODUCTION = "yes" Then
      fn_GetSSLHREF = HREF_SSL & strPage
   Else
      fn_GetSSLHREF = strPage
   End If

End Function


'Sub Display Header
'*********************************************
Sub DisplayHeader(strCS, strATC)

%>
<HTML>
<HEAD>
<TITLE>User Page</TITLE>
<!-- <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"> -->
<!-- <meta http-equiv="Content-Type" content="text/html; charset=windows-1252"> -->
<meta content="text/html; charset=unicode" http-equiv="Content-Type">
<meta name="Pragma" CONTENT="no-cache">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name=VI60_DTCScriptingPlatform content="Server (ASP)">
<meta name=VI60_defaultClientScript content=JavaScript>
</HEAD>

<BODY>
<TABLE cellSpacing=0 cellPadding=0 border=0>
  <TR>
  <TD align=left valign=top>
  <TABLE cellSpacing=0 cellPadding=0 border=0>
          <TR valign=center><TD width=800 align=left><a href="<% = fn_CreateURL(PAGE_LOGOUT,strCs,ACTION_LOGOUT_GOHOME,"") %>"><img align=left valign=center src="Pics\PS_logo1.gif" border=0></a></TD></TR>
  </TABLE>
  </TD>
  </TR> 
  <TR>
   <TD><HR color="#0066CC" width=800></TD>
  </TR>
   
  <TR>   
  <TD align=left valign=top>  

  <!-- ************************* -->
  <TABLE cellSpacing=0 cellPadding=0 border=0> 
   <TR>
     <TD align=left valign=top width=165>
       <%
         DisplayNavLinks strCs, strATC
       %>
     </TD>
     
     <TD align=left valign=top>     
<% 
End Sub   'DisplayHeader


'Sub Display Footer
'*********************************************
Sub DisplayFooter(strCS)
%>
     </TD>


    </TR>
  </TABLE>

  <!-- ************************* -->
  </TD>
  </TR>
  
  <TR>
   <TD><HR color="#0066CC" width=800></TD>
  </TR>
   <TR>
  <TD><center><FONT size=2><BR>2002 PasswordSupport.com &#8482;.</FONT></center></td>
  </TR>

</TABLE>
</BODY>

</HTML>
<% 
End Sub  'Display Footer


' Display Navigation links
'*****************************************************************************
Sub DisplayNavLinks(strC, strATC)
Dim blnAllow
Dim arryTmp
Dim i

blnAllow = False
blnAllow = fn_GetAccountTypeFullAccessAllowed(strATC)

'If GLOB_PRODUCTION = "yes" Then

    If Not blnAllow Then
        Response.Write "<a href=""" & fn_CreateURL(PAGE_LOGOUT,strC,ACTION_LOGOUT_GOHOME,"") & """><font face=""Arial"" size=-1>PasswordSupport Home</font></a> <BR><BR>"
        Response.Write "<a href=""" & fn_CreateURL(PAGE_LOGOUT,strC,ACTION_LOGOUT_NORMAL,"") & """><font face=""Arial"">LogOut</font></a> <BR><BR>"         
        Response.Write "<a href=""" & fn_CreateSSLURL(PAGE_USERMAIN,strC,ACTION_MAIN,"") & """><font face=""Arial"" size=-2>User Main Menu</font></a> <BR>" & fn_InsertSpaces(4) 
        Response.Write "<a href=""" & fn_CreateSSLURL(PAGE_DATAPASSWORD,strC,ACTION_MAIN,"") & """><font face=""Arial"" size=-2>My Passwords</font></a> <BR>" & fn_InsertSpaces(4)
        Response.Write "<font face=""Arial"" size=-2 color=""GRAY"">My Calendar</font> <BR>" & fn_InsertSpaces(4)
        Response.Write "<font face=""Arial"" size=-2 color=""GRAY"">Friends & Family Contacts</font> <BR>" & fn_InsertSpaces(4) 
        Response.Write "<font face=""Arial"" size=-2 color=""GRAY"">Business Contacts</font> <BR>" & fn_InsertSpaces(4) 
        Response.Write "<font face=""Arial"" size=-2 color=""GRAY"">My Addresses</font> <BR>" & fn_InsertSpaces(4) 
        Response.Write "<font face=""Arial"" size=-2 color=""GRAY"">My NoteBook</font> <BR>" & fn_InsertSpaces(4) 
        
    Else
        Response.Write "<a href=""" & fn_CreateURL(PAGE_LOGOUT,strC,ACTION_LOGOUT_GOHOME,"") & """><font face=""Arial"" size=-1>PasswordSupport Home</font></a> <BR><BR>"
        Response.Write "<a href=""" & fn_CreateURL(PAGE_LOGOUT,strC,ACTION_LOGOUT_NORMAL,"") & """><font face=""Arial"">LogOut</font></a> <BR><BR>"         
        Response.Write "<a href=""" & fn_CreateSSLURL(PAGE_USERMAIN,strC,ACTION_MAIN,"") & """><font face=""Arial"" size=-2>User Main Menu</font></a> <BR>" & fn_InsertSpaces(4) 
        Response.Write "<a href=""" & fn_CreateSSLURL(PAGE_DATAPASSWORD,strC,ACTION_MAIN,"") & """><font face=""Arial"" size=-2>My Passwords</font></a> <BR>" & fn_InsertSpaces(4)
        Response.Write "<a href=""" & fn_CreateSSLURL(PAGE_DATACALENDAR,strC,ACTION_MAIN,"") & """><font face=""Arial"" size=-2>My Calendar</font></a> <BR>" & fn_InsertSpaces(4)
        Response.Write "<a href=""" & fn_CreateSSLURL(PAGE_DATAFRIEND,strC,ACTION_MAIN,"") & """><font face=""Arial"" size=-2>Friends & Family Contacts</font></a> <BR>" & fn_InsertSpaces(4) 
        Response.Write "<a href=""" & fn_CreateSSLURL(PAGE_DATABUSINESS,strC,ACTION_MAIN,"") & """><font face=""Arial"" size=-2>Business Contacts</font></a> <BR>" & fn_InsertSpaces(4) 
        Response.Write "<a href=""" & fn_CreateSSLURL(PAGE_DATAMYADDRESS,strC,ACTION_MAIN,"") & """><font face=""Arial"" size=-2>My Addresses</font></a> <BR>" & fn_InsertSpaces(4) 
        Response.Write "<a href=""" & fn_CreateSSLURL(PAGE_DATANOTE,strC,ACTION_MAIN,"") & """><font face=""Arial"" size=-2>My NoteBook</font></a> <BR>" & fn_InsertSpaces(4)     
    End If

'Else

 '   If Not blnAllow Then
 '   Response.Write "<a href=""" & fn_CreateURL(PAGE_LOGOUT,strC,ACTION_LOGOUT_GOHOME,"") & """>PasswordSupport Home</a> <BR><BR>"
 '   Response.Write "<a href=""" & fn_CreateURL(PAGE_LOGOUT,strC,ACTION_LOGOUT_NORMAL,"") & """>LogOut</a> <BR><BR>"         
 '   Response.Write "<a href=""" & fn_CreateURL(PAGE_USERMAIN,strC,ACTION_MAIN,"") & """>Main Menu</a> <BR>" & fn_InsertSpaces(4) 
 '       Response.Write "<a href=""" & fn_CreateURL(PAGE_DATAFRIEND,strC,ACTION_MAIN,"") & """>F & F Contacts</a> <BR>" & fn_InsertSpaces(4) 
 '       Response.Write "<font color=""GRAY"">Business Contacts</font> <BR>" & fn_InsertSpaces(4) 
 '       Response.Write "<a href=""" & fn_CreateURL(PAGE_DATAPASSWORD,strC,ACTION_MAIN,"") & """>My Passwords</a> <BR>" & fn_InsertSpaces(4)
 '       Response.Write "<font color=""GRAY"">My Addresses</font> <BR>" & fn_InsertSpaces(4) 
 '       Response.Write "<font color=""GRAY"">My NoteBook</font> <BR>" & fn_InsertSpaces(4) 
 '       Response.Write "<font color=""GRAY"">My Calendar</font> <BR>"
 '   Else
 '   Response.Write "<a href=""" & fn_CreateURL(PAGE_LOGOUT,strC,ACTION_LOGOUT_GOHOME,"") & """>PasswordSupport Home</a> <BR><BR>"
 '   Response.Write "<a href=""" & fn_CreateURL(PAGE_LOGOUT,strC,ACTION_LOGOUT_NORMAL,"") & """>LogOut</a> <BR><BR>"         
 '   Response.Write "<a href=""" & fn_CreateURL(PAGE_USERMAIN,strC,ACTION_MAIN,"") & """>Main Menu</a> <BR>" & fn_InsertSpaces(4) 
 '       Response.Write "<a href=""" & fn_CreateURL(PAGE_DATAFRIEND,strC,ACTION_MAIN,"") & """>F & F Contacts</a> <BR>" & fn_InsertSpaces(4) 
 '       Response.Write "<a href=""" & fn_CreateURL(PAGE_DATABUSINESS,strC,ACTION_MAIN,"") & """>Business Contacts</a> <BR>" & fn_InsertSpaces(4) 
 '       Response.Write "<a href=""" & fn_CreateURL(PAGE_DATAPASSWORD,strC,ACTION_MAIN,"") & """>My Passwords</a> <BR>" & fn_InsertSpaces(4)
 '       Response.Write "<a href=""" & fn_CreateURL(PAGE_DATAMYADDRESS,strC,ACTION_MAIN,"") & """>My Addresses</a> <BR>" & fn_InsertSpaces(4) 
 '       Response.Write "<a href=""" & fn_CreateURL(PAGE_DATANOTE,strC,ACTION_MAIN,"") & """>My NoteBook</a> <BR>" & fn_InsertSpaces(4)
 '       Response.Write "<a href=""" & fn_CreateURL(PAGE_DATACALENDAR,strC,ACTION_MAIN,"") & """>My Calendar</a> <BR>" 
 '   End If

'End If


End Sub


' This Sub checks the last page cookie and if it is a valid referring 
' page for the current page then the current page becomes the new last page   
'*******************************************************************************   
Sub sub_CheckLastPage(strPage) 
   
   Dim strLP
   
   strLP = Session("sess_strLastPage")
   Session("sess_strLastPage") = strPage
   
   Select Case strLP
      
      Case "":  ' There is no last page
         Call sub_HandlePSError(ERR_INVALID_LAST_PAGE) 
                    
      Case Else: ' there is a last page
   
   End Select

End Sub


' Create a URL string
'*******************************************************
Function fn_CreateURL(strPage, strCs, strAct, strOpt)
  
  Dim strTmp
  
  If strOpt <> "" Then
    'strTmp = strCs & "|ACT:=" & strAct & "|NOW:=" & fn_GetNowAsYYYYMMDD_HHMMSS & "|OPT:=" & strOpt & "|"
    strTmp = strCs & "|ACT:=" & strAct & "|OPT:=" & strOpt & "|"
  Else
    'strTmp = strCs & "|ACT:=" & strAct & "|NOW:=" & fn_GetNowAsYYYYMMDD_HHMMSS & "|"
    strTmp = strCs & "|ACT:=" & strAct & "|"
  End If
 
  strTmp = fn_EncryptString(strTmp) 
  
  strTmp = Server.URLEncode(strTmp)
  fn_CreateURL = strPage & "?x1=" & strTmp 
  
End Function


' Create a SSL URL string
'*******************************************************
Function fn_CreateSSLURL(strPage, strCs, strAct, strOpt)
  
  Dim strTmp
  
  If strOpt <> "" Then
    'strTmp = strCs & "|ACT:=" & strAct & "|NOW:=" & fn_GetNowAsYYYYMMDD_HHMMSS & "|OPT:=" & strOpt & "|"
    strTmp = strCs & "|ACT:=" & strAct & "|OPT:=" & strOpt & "|"
  Else
    'strTmp = strCs & "|ACT:=" & strAct & "|NOW:=" & fn_GetNowAsYYYYMMDD_HHMMSS & "|"
    strTmp = strCs & "|ACT:=" & strAct & "|"
  End If
 
  strTmp = fn_EncryptString(strTmp) 
  
  strTmp = Server.URLEncode(strTmp)
  
  
  fn_CreateSSLURL = fn_GetSSLHREF(strPage) & "?x1=" & strTmp 

  
End Function


' Get an Encrypted Query String Variable
'*******************************************************
Function fn_GetEncryptQryStringVar(ByVal strVarName)
   
   Dim strTmp
   strTmp = Request(strVarName)

   If strTmp <> "" Then
      strTmp = fn_DecryptString(strTmp)
   End If
   
   fn_GetEncryptQryStringVar = strTmp
   
End Function


' Create a Security Cookie
'*******************************************************
Function fn_CreateSecurityCookie(strUN, intAID)
       Dim strTmp2  'As String
       Dim strC
       
       strC = fn_GetRandomAlphaNumeric(5,10)
       strTmp2 = "SID:=" & Session.SessionID & "|NOW:=" & fn_GetNowAsYYYYMMDD_HHMMSS(Now()) & "|UID:=" & strUN & "|AID:=" & intAID & "|"
       sub_CreateEncryptCookie strC, strTmp2
       
       fn_CreateSecurityCookie = strC

End Function       


' Create an Encrypted Cookie
'*******************************************************
Sub sub_CreateEncryptCookie(ByVal strName, ByVal strValue)
   
   If strValue <> "" Then
       strValue = fn_EncryptString(strValue) 
   End If
   Response.Cookies(strName) = strValue    

End Sub   

' Get Encrypted Cookie Data
'*******************************************************
Function fn_GetEncryptCookie(strName)
  
   Dim strValue
   strValue = Request.Cookies(strName)
   If strValue <> "" Then
       strValue = fn_DecryptString(strValue) 
   End If
     
   fn_GetEncryptCookie = strValue 

End Function 


' Create a Simple Cookie
'*******************************************************
Sub sub_CreateCookie(strName, strValue)

   Response.Cookies(strName) = strValue      

End Sub  

' Get a normal Cookie Data
'*******************************************************
Function fn_GetCookie(strName)

   fn_GetCookie = Request.Cookies(strName)

End Function  


' Create an Encrypted Key in a cookie (multi-value cookie)
'*******************************************************
Sub sub_CreateEnCryptCookieKey(ByVal strCN, ByVal strKN, ByVal strVal)

   If strVal <> "" Then
       strVal = fn_EncryptString(strVal) 
   End If
   Response.Cookies(strCN)(strKN) = strVal      
  
End Sub   

'Get an Encrypted Key from a Cookie (multi-value cookie)
'*******************************************************
Function fn_GetEncryptCookieKey(ByVal strCN, ByVal strKN)

   Dim strVal
   strVal = Request.Cookies(strCN)(strKN)  
   If strVal <> "" Then
       strVal = fn_DecryptString(strVal) 
   End If
     
   fn_GetEncryptCookieKey = strVal 
  
End Function   


' Create a Key in a cookie (multi-value cookie)
'*******************************************************
Sub sub_CreateCookieKey(ByVal strCN, ByVal strKN, ByVal strVal)

   Response.Cookies(strCN)(strKN) = strVal      
  
End Sub   

'Get a Key from a Cookie (multi-value cookie)
'*******************************************************
Function fn_GetCookieKey(ByVal strCN, ByVal strKN)

   fn_GetCookieKey = Request.Cookies(strCN)(strKN)  
  
End Function  


'Function fn_GetNowAsYYYYMMDD_HHMMSS
'*********************************************
Function fn_GetNowAsYYYYMMDD_HHMMSS(dtDateTime)
  fn_GetNowAsYYYYMMDD_HHMMSS = Year(dtDateTime) & "-" & Right("0" & Month(dtDateTime),2) & "-" & Right("0" & day(dtDateTime),2) & " " _
                  & Right("0" & Hour(dtDateTime),2) & ":" & Right("0" & Minute(dtDateTime),2) & ":" & Right("0" & Second(dtDateTime),2)
End Function


'Function fn_GetDateAsYYYYMMDD
'*********************************************
Function fn_GetDateAsYYYYMMDD(dtDate)
   fn_GetDateAsYYYYMMDD = Year(dtDate) & "-" & Right("0" & Month(dtDate),2) & "-" & Right("0" & Day(dtDate),2)
End Function


'Function fn_GetTimeAsHHMMSS
'*********************************************
Function fn_GetTimeAsHHMMSS()
  Dim dt1 'As DateTime
  dt1 = Time()
  fn_GetTimeAsHHMMSS = Right("0" & Hour(dt1),2) & ":" & Right("0" & Minute(dt1),2) & ":" & Right("0" & Second(dt1),2)
End Function


'Function fn_CheckAccountTypeFullAccess
'*********************************************
Function fn_CheckAccountTypeFullAccess(intType)
   
   fn_CheckAccountTypeFullAccess = 0
   Dim arryTmp
   arryTmp = Application("app_arryAcctType")
  
   Dim i
   For i = 0 to Ubound(arryTmp, 2)
      If CStr(arryTmp(0,i)) = Cstr(intType) Then
         If arryTmp(2,i) = True Then
            fn_CheckAccountTypeFullAccess = 1
         End If
      End If 
   Next
   
   Erase arryTmp
  
End Function


'Function fn_GetAccountTypePayingCustomer
'*********************************************
Function fn_GetAccountTypePayingCustomer(strType)
   
   fn_GetAccountTypePayingCustomer = False
   Dim arryTmp
   arryTmp = Application("app_arryAcctType")
  
   Dim i
   For i = 0 to Ubound(arryTmp, 2)
      If arryTmp(0,i) = strType Then
         fn_GetAccountTypePayingCustomer = arryTmp(3,i)
         Erase arryTmp
         Exit Function
      End If 
   Next
   
   Erase arryTmp
  
End Function


'Function fn_GetAccountTypeFullAccessAllowed
'*********************************************
Function fn_GetAccountTypeFullAccessAllowed(strType)
   
   fn_GetAccountTypeFullAccessAllowed = False
   Dim arryTmp
   arryTmp = Application("app_arryAcctType")
  
   Dim i
   For i = 0 to Ubound(arryTmp, 2)
      If arryTmp(0,i) = strType Then
         fn_GetAccountTypeFullAccessAllowed = arryTmp(2,i)
         Erase arryTmp 
         Exit Function
      End If 
   Next
   
   Erase arryTmp 
   
End Function


'Function fn_GetAccountTypeDesc
'*********************************************
Function fn_GetAccountTypeDesc(strType)
   
   fn_GetAccountTypeDesc = 0
   Dim arryTmp
   arryTmp = Application("app_arryAcctType")
  
   Dim i
   For i = 0 to Ubound(arryTmp, 2)
      If CStr(arryTmp(0,i)) = strType Then
         fn_GetAccountTypeDesc = arryTmp(1,i)
      End If 
   Next
   
   Erase arryTmp
  
End Function


'Function fn_GetStateCode
'*********************************************
Function fn_GetStateCode(intID)
   
   fn_GetStateCode = 0
   Dim arryTmp
   arryTmp = Application("app_arryStates")
  
   Dim i
   For i = 0 to Ubound(arryTmp, 2)
      If CStr(arryTmp(0,i)) = Cstr(intID) Then
         fn_GetStateCode = arryTmp(1,i)
         Exit Function
      End If 
   Next
   
   Erase arryTmp
  
End Function


'Function fn_GetCCTypeDesc
'*********************************************
Function fn_GetCCTypeDesc(intType)
   
   fn_GetCCTypeDesc = 0
   Dim arryTmp
   arryTmp = Application("app_arryCreditCardType")
  
   Dim i
   For i = 0 to Ubound(arryTmp, 2)
      If Cstr(arryTmp(0,i)) = CStr(intType) Then
         fn_GetCCTypeDesc = arryTmp(1,i)
         Exit Function
      End If 
   Next
   
   Erase arryTmp
  
End Function


'  Remove any unneccassary Session Variables
'********************************************************
Sub sub_RemoveSessionVariables

     Dim arryTmp
     Dim strTmp
     
     arryTmp = Session("sess_arryCSE")
     If IsArray(arryTmp) Then
        Session.Contents.Remove("sess_arryCSE")
        Erase arryTmp
     End If
     
     
     arryTmp = Session("sess_arryMonthInfo")
     If IsArray(arryTmp) Then
        Session.Contents.Remove("sess_arryMonthInfo")
        Erase arryTmp
     End If
     
     
     arryTmp = Session("sess_arryDayInfo")
     If IsArray(arryTmp) Then
        Session.Contents.Remove("sess_arryDayInfo")
        Erase arryTmp
     End If
     
     'Session("sess_arryUserData") 
     
     strTmp = Session("sess_dtCurDate")
     If strTmp <> "" Then
        Session.Contents.Remove("sess_dtCurDate")
     End If

End Sub



' This sub checks the last page and checks the login and parses the encrypted query string
' This sub returns the strCs, the query string, and the action
'**********************************************************************8
Sub sub_MechaCheck(ByVal strPage, ByVal strQSVar, strCs, strQry, strAct)
    
    Dim intRC
    
    Call sub_CheckLastPage(strPage)
    strQry = fn_GetEncryptQryStringVar(strQSVar)

    strCs = Left(strQry,InStr(strQry,"|")-1)
    intRC = fn_CheckForLogin(strCs)
    If intRC <> 0 Then
       Call sub_HandlePSError(ERR_NO_LOGIN)
    End If    
    
    Select Case strPage
       Case PAGE_DATACALENDAR:
          If fn_GetAccountTypeFullAccessAllowed(Session("sess_strATC")) <> True Then
             Call sub_HandlePSError(ERR_ACCESS_VIOLATION)
          End If 
          
       Case PAGE_DATABUSINESS, PAGE_DATAMYADDRESS, PAGE_DATANOTE:
          sub_RemoveSessionVariables
          If fn_GetAccountTypeFullAccessAllowed(Session("sess_strATC")) <> True Then
             Call sub_HandlePSError(ERR_ACCESS_VIOLATION)
          End If 
               
       Case Else:   
         sub_RemoveSessionVariables
    
    End Select
    
    strAct = fn_GetItemFromString(strQry, "ACT:=", "|")
  
End Sub


'Sub Check For Login
'*********************************************
Function fn_CheckForLogin(strCs)

   fn_CheckForLogin = -1

   If Session("sess_blnLoggedIn") = True and Session("sess_strCs") = strCs Then
       fn_CheckForLogin = 0
   End If

End Function


' Extract an item from a String
'******************************************************
'*******************************************************
Function fn_GetItemFromString(strIn, strKey, strDelim)

   Dim intLocKey
   Dim intLocDelim
  
   intLocKey = InStr(strIn,strKey)
   If intLocKey = 0 Then
      fn_GetItemFromString = ""
      Exit Function
   End If

   intLocDelim = InStr(intLocKey,strIn,strDelim)

   fn_GetItemFromString = Mid(strIn, intLocKey + Len(strKey), intLocDelim - (intLocKey + Len(strKey)))
  
End Function



' this function gets a HEX or HTML color code based on a color_id
'*****************************************************
Function fn_GetColorCode(ByVal intIndex)

   Dim arryTmp
   arryTmp = Application("app_arryDCColors")
   If intIndex < 0 or intIndex > UBound(arryTmp, 2) Then
      fn_GetColorCode = "BLACK"
   Else
      fn_GetColorCode = arryTmp(1,intIndex)
   End If

   Erase arryTmp
   
End Function

' this function gets a color Name based on a color_id
'*****************************************************
Function fn_GetColorName(ByVal intIndex)

   Dim arryTmp
   arryTmp = Application("app_arryDCColors")
   If intIndex < 0 or intIndex > UBound(arryTmp,2) Then
      fn_GetColorName = "Black"
   Else
      fn_GetColorName = arryTmp(2,intIndex)
   End If

   Erase arryTmp
   
End Function


' Inserts Spaces
'******************************************************
'*******************************************************
Function fn_InsertSpaces(intNumSpaces)
   fn_InsertSpaces = ""
   
   If intNumSpaces = 0 or Not IsNumeric(intNumSpaces) Then
      Exit Function
   End If

   Dim i
   For i = 1 to intNumSpaces
      fn_InsertSpaces = fn_InsertSpaces & "&nbsp;"
   Next

End Function



' Unescape Delims  - Replace any escaped delim chars with the actual chars
'******************************************************
'*******************************************************
Function fn_UnescapeDelims(ByVal strIn)
   
   strIn = Replace(strIn,ESC_TILDY,"~")
   strIn = Replace(strIn,ESC_PIPE,"|")
   strIn = Replace(strIn,ESC_CARROT,"^")
   fn_UnescapeDelims = strIn

End Function

' Escape Delims  - Replace any delim chars with the escape chars
'******************************************************
'*******************************************************
Function fn_EscapeDelims(ByVal strIn)
   
   strIn = Replace(strIn,"~",ESC_TILDY)
   strIn = Replace(strIn,"|",ESC_PIPE)
   strIn = Replace(strIn,"^",ESC_CARROT)
   fn_EscapeDelims = strIn

End Function


' Unescape SQL  - Replace any escaped single quotes with an actual single quote
'******************************************************
'*******************************************************
Function fn_UnescapeSQL(strIn)
  
   strIn = Replace(strIn,ESC_SQ,"'")
   fn_UnescapeSQL = strIn

End Function

' Escape SQL  - Replace single quotes with the escape char
'******************************************************
'*******************************************************
Function fn_EscapeSQL(ByVal strIn)
   
   strIn = Replace(strIn,"'",ESC_SQ)
   fn_EscapeSQL = strIn

End Function




' Catch Asp Error and Transfer to ASP Error Page
'******************************************************
'*******************************************************
Sub sub_ErrorCatch()

    If Err.Number <> 0 Then
       'Response.Write Err.Number
       'Response.End
       Server.Transfer PAGE_500100ERROR
    End If
  
End Sub


' Handle a PS Error and Transfer to PSError Page
'******************************************************
'*******************************************************
Sub sub_HandlePSError(intErr)

    Session("sess_intPSError") = intErr
    Server.Transfer PAGE_PSERROR
  
End Sub


' Sub  Calculate Page Info
'******************************************************
'*******************************************************
Sub  sub_CalculatePagesInfo (ByVal intUB, intStart, intEnd, intPage, intNewNumPages)

    intStart = 0
    intEnd = MAX_DESC_PER_PAGE - 1
    intNewNumPages = 1
    
    If (intUB + 1) > MAX_DESC_PER_PAGE Then
      Dim dblNumPages
      Dim intFix
      
       dblNumPages = CDbl(intUB + 1) / CDbl(MAX_DESC_PER_PAGE) 
       intFix = Fix(dblNumPages)
       If dblNumPages - intFix > 0 Then
          intNewNumPages = intFix + 1
       Else
          intNewNumPages = intFix
       End If
       
       sub_CreateCookie COOK_NEWNUMPAGES, CStr(intNewNumPages)
       
       If intPage > intNewNumPages Then
          intPage = intNewNumPages
       End If
       
       sub_CreateCookie COOK_OLDPAGENUM, CStr(intPage)
    
       If intPage > 1 Then
          intStart = ((intPage - 1) * MAX_DESC_PER_PAGE)     
          intEnd = intStart + MAX_DESC_PER_PAGE - 1
       End If
    
    End If
    
    If intEnd > intUB Then
       intEnd = intUB
    End If
    
End Sub


' Sub Display Page Number List
'**************************************************************
'**************************************************************
Sub sub_DisplayPageNumberList (ByVal strCs, ByVal strPage, ByVal strAction, ByVal intPage, ByVal intNewNumPages)
   
   If intNewNumPages > 1 Then
        %>
        <FORM action="<% = fn_CreateURL(strPage,strCs,strAction,"") %>" method="post">
        <input type="hidden" name="txtOldPage" value="<% = CStr(intPage) %>">
        <input type="hidden" name="txtOldNumPages" value="<% = CStr(intNewNumPages) %>">
        Page <% = CStr(intPage)%> of <% = CStr(intNewNumPages) & fn_InsertSPaces(4) %><input type="submit" name="cmdSubmit" value="<<"><% = fn_InsertSpaces(3) %><select name="lstPageNumber" >
            <%
              Dim i
              For i = 1 to intNewNumPages
                If i = intPage Then
                   Response.Write "<option value=""" & CStr(i) & """ selected>Page " & CStr(i)
                Else
                   Response.Write "<option value=""" & CStr(i) & """>Page " & CStr(i)
                End If
              Next
            %>     
        </select><% = fn_InsertSpaces(1) %><input type="submit" name="cmdSubmit" value="Go"><% = fn_InsertSpaces(3) %><input type="submit" name="cmdSubmit" value=">>">  
        </FORM>
        <%
    End If

End Sub


' Sub Validate Page Select
'**************************************************************
'**************************************************************
Sub sub_ValidatePageSelect(ByVal strCs, ByVal strPage, ByVal strAction)     
    Dim strSubmit
    Dim intPage
    Dim intLastPage
    Dim intNumPages
    
    
    ' Get the submit value
    '***********************************
    strSubmit = Request("cmdSubmit")
    
    ' Get the old Page value
    '***********************************
    intLastPage = Request("txtOldPage")
    If intLastPage = "" or Not IsNumeric(intLastPage) Then 
       intLastPage = 1
    Else
       intLastPage = CInt(intLastPage)
    End If
    
    ' Get the old number of pages
    '***********************************
    intNumPages = Request("txtOldNumPages")
    If intNumPages = "" or Not IsNumeric(intNumPages) Then 
       intNumPages = 1
    Else
       intNumPages = CInt(intNumPages)
    End If

   '
   '***********************************
   Select Case strSubmit

       Case ">>":
          If intLastPage < intNumPages Then
             intPage = intLastPage + 1
          Else
             intPage = intNumPages   
          End If
          
       Case "<<":
          If intLastPage > 1 Then
             intPage = intLastPage - 1
          Else
             intPage = 1   
          End If
          
       Case "Go":
          intPage = Request("lstPageNumber")
          If intPage = "" or Not IsNumeric(intPage) Then 
             intPage = 1
          Else
             intPage = CInt(intPage)
          End If
          
       Case Else:     
          intPage = 1
      
   End Select   

   Response.Redirect fn_CreateURL(strPage,strCs,strAction,intPage)

End Sub



'
'*********************************************************
Sub sub_SetApplicationVariables()
   Dim intRC
   Dim arryTmp

   Application.Lock
   Application("app_blnSet") = 0
   Application.Unlock
   
   
   ' Set Time Zone Array
   '***************************************************
   intRC = fn_DL_GetTimeZones(arryTmp)
   If intRC <> 0 Then
      Call sub_HandlePSError(ERR_GET_TIMEZONE)
   End If
   
   Application.Lock
   Application("app_arryTZ") = arryTmp
   Application.Unlock
   Erase arryTmp
   
   
   ' Set Account Statuses Array
   '**************************************************
   intRC = fn_DL_GetAccountStatuses(arryTmp)
   If intRC <> 0 Then
      Call sub_HandlePSError(ERR_GET_ACCTSTATUSES)
   End If
   Application.Lock
   Application("app_arryAcctStat") = arryTmp
   Application.Unlock
   Erase arryTmp
   
      ' Set Account Types Array
   '**************************************************
   intRC = fn_DL_GetAccountTypes(arryTmp)
   If intRC <> 0 Then
      Call sub_HandlePSError(ERR_GET_ACCTTYPES)
   End If
   Application.Lock
   Application("app_arryAcctType") = arryTmp
   Application.Unlock
   Erase arryTmp
   
   ' Set States Array
   '**************************************************
   intRC = fn_DL_GetStates(arryTmp)
   If intRC <> 0 Then
      Call sub_HandlePSError(ERR_GET_STATES)
   End If
   Application.Lock
   Application("app_arryStates") = arryTmp
   Application.Unlock
   Erase arryTmp
   
   ' Set Countries Array
   '**************************************************
   intRC = fn_DL_GetCountries(arryTmp)
   If intRC <> 0 Then
      Call sub_HandlePSError(ERR_GET_COUNTRIES)
   End If
   Application.Lock
   Application("app_arryCountries") = arryTmp
   Application.Unlock
   Erase arryTmp


   ' Set Credit Card Types Array
   '**************************************************
   intRC = fn_DL_GetCreditCardTypes(arryTmp)
   If intRC <> 0 Then
      Call sub_HandlePSError(ERR_GET_CCTYPE)
   End If
   Application.Lock
   Application("app_arryCreditCardType") = arryTmp
   Application.Unlock
   Erase arryTmp



   ' Set DataCalendarColors Array
   '**************************************************
   intRC = fn_DL_GetDataCalendarColors(arryTmp)
   If intRC <> 0 Then
      Call sub_HandlePSError(ERR_GET_DCCOLOR)
   End If
   Application.Lock
   Application("app_arryDCColors") = arryTmp
   Application.Unlock
   Erase arryTmp
   
  
   'If IsArray(arryTmp) Then
   '  Response.Write "Is Array"
   'Else
   '   Response.Write "Is Not Array"
   'End If
   'Response.End


   'arryTmp = Application("app_arryAcctStat")
   'Dim i
   'For i = 0 to 5
   '   Response.Write  arryTmp(0,i) & " - " & arryTmp(1,i) & "<BR>"
   'Next
   'Erase arryTmp
   'Response.End
   
   Application.Lock
   Application("app_blnSet") = 1
   Application.Unlock
   
End Sub


%>





















