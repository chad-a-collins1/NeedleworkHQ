<!--#INCLUDE FILE="clsUpload.asp"-->
<%
Dim oUpload
Dim oFile
Dim sFileName
Dim oFSO
Dim sPath
Dim sNewData
Dim nLength
Dim oFile2
Dim bytBinaryData

Const nForReading = 1
Const nForWriting = 2
Const nForAppending = 8

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("./") & "\db\NWHQ.mdb"

Set oRs2 = Server.CreateObject("ADODB.Recordset")
sSQL2 = "SELECT Top 1 pID From tblConsignPatterns Order By pID Desc"
oRs2.Open sSQL2, oConn

pID = oRs2.Fields("pID")

oRs2.Close
Set oRs2 = Nothing

' grab the uploaded file data
Set oUpload = New clsUpload
Set oFile = oUpload("File1")

' parse the file name
sFileName = pID & Right((oFile.FileName), 4)
sFileName_original = oFile.FileName
If Not InStr(sFileName, "\") = 0 Then
	sFileName = Mid(sFileName, InStrRev(sFileName, "\") + 1)
End If

If Not InStr(sFileName_Original, "\") = 0 Then
	sFileName_Original = Mid(sFileName_Original, InStrRev(sFileName_Original, "\") + 1)
End If

' Convert the binary data to Ascii
bytBinaryData = oFile.BinaryData
nLength = LenB(bytBinaryData)
For nIndex = 1 To nLength
	sNewData = sNewData & Chr(AscB(MidB(bytBinaryData, nIndex, 1)))
Next

' Save the file to the file system
sPath = Server.MapPath(".\dev\ConsignmentShop") & "\"
Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
oFSO.OpenTextFile(sPath & sFileName, nForWriting, True).Write sNewData
Set oFSO = Nothing


Set oFile1 = oFile


Set oFile = Nothing
Set oUpload = Nothing

'********************************************************************************


Set oRs = Server.CreateObject("ADODB.Recordset")

sSQL = "SELECT FileID, FileName, Name, FileSize, pID, ContentType, BinaryData FROM tblConsignmentFiles WHERE 1=2"

oRs.Open sSQL, oConn, 3, 3

If oFile1.ContentType = "image/pjpeg" Then

		oRs.AddNew
		oRs.Fields("FileName") = sFileName_original 
		oRs.Fields("Name") = sFileName
		oRs.Fields("pID") = pID
		oRs.Fields("FileSize") = oFile1.Length
		oRs.Fields("ContentType") = oFile1.ContentType
		oRs.Fields("BinaryData").AppendChunk = oFile1.BinaryData & ChrB(0)
		oRs.Update

		oRs.Close
		oConn.Close

Else
		oRs.Close
		oConn.Close
		Response.Write "<center><h3><font color=darkred>The File you uploaded was not a .JPG, please go back and try again.</font></h3></center>"
End If		
		

Set oConn = Nothing
Set oFile1 = Nothing

'*************************************************************************************************

%>



<BODY background="paper_old.gif">
<STYLE type=text/css>
	p {font-size: 8pt;font-family: "Verdana";color: "darkblue"; }
	td {font-size: 8pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 8pt;font-family: "Verdana"; color: "black";}
	a {font-size: 8pt;font-family: "Verdana"; color: "darkblue";}
</STYLE>
<BR><BR><BR><center>

<Table width="65%" background="yellow.jpg">
<TR>
<TH>
<CENTER>
<B>Upload Your PC Stitch Pattern (.PAT) File</B>
</CENTER><BR><BR>
</TH>
</TR>
<TR>
<TD bgcolor="lightyellow"><font color="red"><b><u>STEP 4:</u></b></font>&nbsp;<b>Upload your PC Stitch (.pat) file. Use the <font color="red">"Browse"</font> feature below to find the file on your computer then click the <font color="red">"Upload (.PAT) FILE"</font> button.</b></TD>
</TR>
<TR><TD><center><BR>
<% 

uid = Request.QueryString("u").Item(1)
pass = Request.QueryString("u").Item(2) 

%>
<FORM method="post" encType="multipart/form-data" action="AddNewConsignPat4.asp?u=<%= uid & "&u=" & pass %>" id=form1 name=form1 onSubmit="return submitForms()">
	<INPUT type="File" name="File1">
	<INPUT type="Submit" value="Upload (.PAT) FILE" id=Submit1 name=Submit1>
</FORM></center></TD></TR>

<TR>
<TD><BR><center>
	
</center></TD>
</TR>
</center>
</Table>
</BODY>

</html>


<SCRIPT Language="Javascript">

//***********This array stores the user name and JPEG entered by the user in the Login text fields below

function submitForms() {
	if (isPAT()) { return true }
	else {return false}
	};


function isPAT() {
if (document.form1.File1.value == "") {
alert ("\n The PC Stitch Pattern field is blank. \n\nPlease upload your PC Stitch Pattern file.")
document.form1.File1.focus();
return false;
}
if ((document.form1.File1.value.indexOf ('.pat',0) == -1)) {
alert ("\n The file you attempted to upload was not a .PAT, please upload a PC Stitch Pattern (.PAT) file only.")
document.form1.File1.select();
document.form1.File1.focus();
return false;
}
return true;
}

</SCRIPT>


