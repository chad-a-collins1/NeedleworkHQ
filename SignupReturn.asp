<% @LANGUAGE = "VBScript"%>
<% Response.Buffer = True %>
<html>
<head>
<script language="javascript1.2">

function fctRedirect()
{

setTimeout("window.navigate('NewMember2.asp?Approved=y')",5000);


}

</script>



</head>
<body onLoad="fctRedirect()">
<br><br><br><br>
<center>Transfering, please wait..............</center>

</body>
</html>


