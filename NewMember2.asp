<% @LANGUAGE = "VBScript" %>
<% Response.Buffer = True %>
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
		rsNewAccount("ApprovedYN") = n
		rsNewAccount("userAlias") = alias
       
        rsNewAccount.Update
        rsNewAccount.MoveFirst 	       
        
        rsNewAccount.Close
        Set rsNewAccount = Nothing
        
        conn.Close
        Set conn = Nothing
%>		

<html>
<head>
<script language="jscript" runat="server">


/*
 * A JavaScript implementation of the Secure Hash Algorithm, SHA-1, as defined
 * in FIPS PUB 180-1
 * Version 2.1-BETA Copyright Paul Johnston 2000 - 2002.
 * Other contributors: Greg Holt, Andrew Kepert, Ydnar, Lostinet
 * Distributed under the BSD License
 * See http://pajhome.org.uk/crypt/md5 for details.
 */

/*
 * Configurable variables. You may need to tweak these to be compatible with
 * the server-side, but the defaults work in most cases.
 */
var hexcase = 0;  /* hex output format. 0 - lowercase; 1 - uppercase        */
var b64pad  = "="; /* base-64 pad character. "=" for strict RFC compliance   */
var chrsz   = 8;  /* bits per input character. 8 - ASCII; 16 - Unicode      */

/*
 * These are the functions you'll usually want to call
 * They take string arguments and return either hex or base-64 encoded strings
 */
function hex_sha1(s){return binb2hex(core_sha1(str2binb(s),s.length * chrsz));}
function b64_sha1(s){return binb2b64(core_sha1(str2binb(s),s.length * chrsz));}
function str_sha1(s){return binb2str(core_sha1(str2binb(s),s.length * chrsz));}
function hex_hmac_sha1(key, data){ return binb2hex(core_hmac_sha1(key, data));}
function b64_hmac_sha1(key, data){ return binb2b64(core_hmac_sha1(key, data));}
function str_hmac_sha1(key, data){ return binb2str(core_hmac_sha1(key, data));}

/*
 * Perform a simple self-test to see if the VM is working
 */
function sha1_vm_test()
{
  return hex_sha1("abc") == "a9993e364706816aba3e25717850c26c9cd0d89d";
}

/*
 * Calculate the SHA-1 of an array of big-endian words, and a bit length
 */
function core_sha1(x, len)
{
  /* append padding */
  x[len >> 5] |= 0x80 << (24 - len % 32);
  x[((len + 64 >> 9) << 4) + 15] = len;

  var w = Array(80);
  var a =  1732584193;
  var b = -271733879;
  var c = -1732584194;
  var d =  271733878;
  var e = -1009589776;

  for(var i = 0; i < x.length; i += 16)
  {
    var olda = a;
    var oldb = b;
    var oldc = c;
    var oldd = d;
    var olde = e;

    for(var j = 0; j < 80; j++)
    {
      if(j < 16) w[j] = x[i + j];
      else w[j] = rol(w[j-3] ^ w[j-8] ^ w[j-14] ^ w[j-16], 1);
      var t = safe_add(safe_add(rol(a, 5), sha1_ft(j, b, c, d)),
                       safe_add(safe_add(e, w[j]), sha1_kt(j)));
      e = d;
      d = c;
      c = rol(b, 30);
      b = a;
      a = t;
    }

    a = safe_add(a, olda);
    b = safe_add(b, oldb);
    c = safe_add(c, oldc);
    d = safe_add(d, oldd);
    e = safe_add(e, olde);
  }
  return Array(a, b, c, d, e);

}

/*
 * Perform the appropriate triplet combination function for the current
 * iteration
 */
function sha1_ft(t, b, c, d)
{
  if(t < 20) return (b & c) | ((~b) & d);
  if(t < 40) return b ^ c ^ d;
  if(t < 60) return (b & c) | (b & d) | (c & d);
  return b ^ c ^ d;
}

/*
 * Determine the appropriate additive constant for the current iteration
 */
function sha1_kt(t)
{
  return (t < 20) ?  1518500249 : (t < 40) ?  1859775393 :
         (t < 60) ? -1894007588 : -899497514;
}

/*
 * Calculate the HMAC-SHA1 of a key and some data
 */
function core_hmac_sha1(key, data)
{
  var bkey = str2binb(key);
  if(bkey.length > 16) bkey = core_sha1(bkey, key.length * chrsz);

  var ipad = Array(16), opad = Array(16);
  for(var i = 0; i < 16; i++)
  {
    ipad[i] = bkey[i] ^ 0x36363636;
    opad[i] = bkey[i] ^ 0x5C5C5C5C;
  }

  var hash = core_sha1(ipad.concat(str2binb(data)), 512 + data.length * chrsz);
  return core_sha1(opad.concat(hash), 512 + 160);
}

/*
 * Add integers, wrapping at 2^32. This uses 16-bit operations internally
 * to work around bugs in some JS interpreters.
 */
function safe_add(x, y)
{
  var lsw = (x & 0xFFFF) + (y & 0xFFFF);
  var msw = (x >> 16) + (y >> 16) + (lsw >> 16);
  return (msw << 16) | (lsw & 0xFFFF);
}

/*
 * Bitwise rotate a 32-bit number to the left.
 */
function rol(num, cnt)
{
  return (num << cnt) | (num >>> (32 - cnt));
}

/*
 * Convert an 8-bit or 16-bit string to an array of big-endian words
 * In 8-bit function, characters >255 have their hi-byte silently ignored.
 */
function str2binb(str)
{
  var bin = Array();
  var mask = (1 << chrsz) - 1;
  for(var i = 0; i < str.length * chrsz; i += chrsz)
    bin[i>>5] |= (str.charCodeAt(i / chrsz) & mask) << (24 - i%32);
  return bin;
}

/*
 * Convert an array of big-endian words to a string
 */
function binb2str(bin)
{
  var str = "";
  var mask = (1 << chrsz) - 1;
  for(var i = 0; i < bin.length * 32; i += chrsz)
    str += String.fromCharCode((bin[i>>5] >>> (24 - i%32)) & mask);
  return str;
}

/*
 * Convert an array of big-endian words to a hex string.
 */
function binb2hex(binarray)
{
  var hex_tab = hexcase ? "0123456789ABCDEF" : "0123456789abcdef";
  var str = "";
  for(var i = 0; i < binarray.length * 4; i++)
  {
    str += hex_tab.charAt((binarray[i>>2] >> ((3 - i%4)*8+4)) & 0xF) +
           hex_tab.charAt((binarray[i>>2] >> ((3 - i%4)*8  )) & 0xF);
  }
  return str;
}

/*
 * Convert an array of big-endian words to a base-64 string
 */
function binb2b64(binarray)
{
  var tab = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  var str = "";
  for(var i = 0; i < binarray.length * 4; i += 3)
  {
    var triplet = (((binarray[i   >> 2] >> 8 * (3 -  i   %4)) & 0xFF) << 16)
                | (((binarray[i+1 >> 2] >> 8 * (3 - (i+1)%4)) & 0xFF) << 8 )
                |  ((binarray[i+2 >> 2] >> 8 * (3 - (i+2)%4)) & 0xFF);
    for(var j = 0; j < 4; j++)
    {
      if(i * 8 + j * 6 > binarray.length * 32) str += b64pad;
      else str += tab.charAt((triplet >> 6*(3-j)) & 0x3F);
    }
  }
  return str;
}

/* End of SHA-1 implementation */
/* *************************** */



/*
 * HOP functions
 * Copyright 2003, CyberSource Corporation.  All rights reserved.
 */

function timestamp()
{
   var d = new Date();
   return( Date.UTC( d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(),
		     d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds(),
		     d.getUTCMilliseconds() ) );
}


function hopHash(data, key)
{
  return b64_hmac_sha1(key, data);
}


/*
 * HOP integration function
 */

function getMerchantID() { return "v7103530"; }
function getPublicKey()  { return "MIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQDLFZgwXPul4qX0yGUzsjq9StFlVpD9Y25luaol1FY+I7gmW1mH6NEJgEMhV0aqJYPFIMIadDE0CRPkiANUPaSRYIoNe9kvK9B1nw7y83M+5HdppJvmYqHBi9UYUmvs62ogQi76JiWq4kO35s9NGZnUOH5m33ByVu5q7v03SKltOwIDAQAB"; }
function getPrivateKey() { return "MIICdwIBADANBgkqhkiG9w0BAQEFAASCAmEwggJdAgEAAoGBAMsVmDBc+6XipfTIZTOyOr1K0WVWkP1jbmW5qiXUVj4juCZbWYfo0QmAQyFXRqolg8Ugwhp0MTQJE+SIA1Q9pJFgig172S8r0HWfDvLzcz7kd2mkm+ZiocGL1RhSa+zraiBCLvomJariQ7fmz00ZmdQ4fmbfcHJW7mru/TdIqW07AgMBAAECgYEAytFfM3W5UJtBGG0GPRHTffZ5l1ZT6Otjdq5s0ej01IxBvfTfPk9ybKWu5V5PUV+z8KxdjaPa+9fRCRwZDwmdDlaQeaDTm5ceZoZH958/kwMLr5DtFEP3f/gpxAMVmzC9X4Wh757XLF8WNHhtDuvrucHpP3nDcEXrSfbUuKE+ddECQQD1S3rVAQB64sqO+xK/4IMLqcLB3QzmJ5uZpB8cyf8BGZNi9dqGn7iHzodvpBCFq9nLnFV5MpxES+XJmL6cTLf5AkEA0/KGNWlIc8FdtbZhQTUQfEk8Tr8QXacl84eu0xWXCx/hc0zU6fDco/0HAsYPowdYK19PmOG8dMWK5C4iK6zj0wJAPistLIMefgaw0+AqdlsOm4whAkVmGYb8VspT4FYJvVugETrCcdBVUoYzqUXpshdGEebDev4qwNyDlr6RwMdo8QJAFC/LPJcUgYHvTPlb9fv53/yRs+ZaxrC+2p0Xt58czcBxlqvAs69vNGdLHaaDosF2Uls3l5YYfv65pdYHByXmZQJBAKlkoaZ7Dop7tQValp4SHCVx7yn7tY5fS+Jqh63U21gYFLkjOFQwrkZDqAkYZM8E+dCmXd2hX6MSLQ3QFUX8cLk="; }

function InsertSignature(amount)
{
  var time = timestamp();
  var merchantID = getMerchantID();
  var data = merchantID + "" + amount + "" + time;
  var pub = getPublicKey();
  var pvt = getPrivateKey();
  var pub_hash = hopHash(data, pub);
  var pvt_hash = hopHash(data, pvt);

  Response.Write('<input type="hidden" name="amount" value="' + amount + '">\n');
  Response.Write('<input type="hidden" name="orderPage_timestamp" value="' + time + '">\n');
  Response.Write('<input type="hidden" name="merchantID" value="' + merchantID + '">\n');
  Response.Write('<input type="hidden" name="orderPage_signaturePublic" value="' + pub_hash + '">\n');
  Response.Write('<input type="hidden" name="orderPage_signaturePrivate" value="' + pvt_hash +'">\n');
}


function VerifySignature(data, signature)
{
	var pub = getPublicKey();
	var pub_hash = hopHash(data, pub);
	return (signature == pub_hash);
}


</SCRIPT></head>
<body onLoad="document.form1.submit();"> 
<form action="https://orderpage.ic3.com/hop/orderform.jsp" method="post" name="form1">


<% InsertSignature("9.99") %>

  <input type="hidden" name="comments" value="Cross Stitch Connection Membership">
  <input type="hidden" name="amount" value="9.99">
  <input type="hidden" name="merchantID" value="v7103530">

<table>
<TR>
<TD size="35"><INPUT type="hidden" Name="billTo_lastName" value="<%= txtLast %>"  size=30 style="HEIGHT: 22px; WIDTH: 180px"></TD>
<TD size="35"><INPUT type="hidden" Name="billTo_firstName" value="<%= txtFirst %>"  size=20 style="HEIGHT: 22px; WIDTH: 180px"></TD>
</TR>

<TR>
<TD size="35"><INPUT type="hidden" Name="billTo_street1" value="<%= txtStreet %>" size=30 style="HEIGHT: 22px; WIDTH: 300px"></TD>
<TD size="35"><INPUT type="hidden" Name="billTo_city" value="<%= txtCity %>" size=30 style="HEIGHT: 22px; WIDTH: 100px"></TD>
</TR>

<TR>
<TD size="35"><INPUT type="hidden" Name="billTo_state" value="<%= txtState %>" size=30 style="HEIGHT: 22px; WIDTH: 180px"></TD>
<TD size="35"><INPUT type="hidden" Name="billTo_postalCode" value="<%= txtPostal %>" size=30 style="HEIGHT: 22px; WIDTH: 80px"></TD>
</TR>

<TR>
<TD size="35"><INPUT type="hidden" Name="txtDesiredID" value="<%= txtDesiredID %>" size=50 style="HEIGHT: 22px; WIDTH: 180px"></TD>
<TD size="35"><INPUT type="hidden" Name="billTo_email" value="<%= txtEmail %>" size=45 style="HEIGHT: 22px; WIDTH: 180px"></TD>
</TR>


<TR>
<TD size="35"><INPUT type="hidden" type="password" Name="txtDesiredPswrd" value="<%= txtDesiredPswrd %>" size=50 style="HEIGHT: 22px; WIDTH: 180px" ></TD>
<TD size="35"><INPUT type="hidden" type="password" Name="txtValidatePswrd" value="<%= txtDesiredPswrd %>" size=50 style="HEIGHT: 22px; WIDTH: 180px"></TD>
</TR>

<TR>
<TD size="35"><Input type="hidden" Name="cboKeywrdType" value="<%= cboKeywrdType %>" style="HEIGHT: 22px; WIDTH: 180px"> </TD>
<TD size="35"><INPUT type="hidden" Name="txtKeywrd" value="<%= txtKeywrd %>" size=50 style="HEIGHT: 22px; WIDTH: 180px"></TD>
</TR>
</table>
</form>
</body>
</html>













